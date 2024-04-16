import mailbox
import email, email.policy, email.utils
import html
import urllib.parse
import chardet
import base64
import struct
import copy
import re
import os
import shutil
import math
import argparse

def flatten( l ):
    for i in l:
        if isinstance( i, list ):
            yield from flatten( i )
        else:
            yield i

# Puts date into consistent format
def format_date( msg ):
    date = msg.get( 'date' )
    try:
        return email.utils.format_datetime(
            email.utils.parsedate_to_datetime( date )
        )
    except ValueError:
        return date

# Helps extracting tricky header info (at times needed for Subject and From)
def get_header_text(msg, item, default='utf-8'):
    header_text = msg.get ( item )
    headers = email.header.decode_header(header_text)

    header_sections = []
    for text, charset in headers:
        if charset is None:
            try:
                encoding = chardet.detect(text)['encoding']
                header_section = text.decode(encoding)
            except:
                header_section = text
        else:
            header_section = text.decode(charset or default)
        if header_section:
            header_sections.append(header_section)
    return ' '.join(header_sections)


def get_payload_text( msg ):
    subtype = msg.get_content_subtype()
    payload = msg.get_payload( decode=True )
    if ( payload is None ): return ''
    charset = msg.get_charset() or chardet.detect( payload )['encoding']
    content = payload.decode( charset )
    if ( subtype == 'plain' ):
        return html.escape( content ).replace( '\n', '<br>' )
    else:
        return content

def payload_get_type( payload, types, type_list ):
    for t in type_list:
        if ( t in types ):
            return get_payload_text( payload[types.index( t )] )
    # TODO: Is this possible? Testing indicates no

def parse_email( msg ):
    content_type = msg.get_content_type()
    maintype = msg.get_content_maintype()
    subtype = msg.get_content_subtype()
    payload = msg.get_payload()
    if ( maintype == 'text' ):
        return [{
            'name': msg.get_filename(),
            'content': get_payload_text( msg ),
            'type': 'text',
        }]
    elif ( maintype == 'multipart' ):
        if ( subtype == 'alternative' ):
            # TODO: Can this contain non-text?
            types = [m.get_content_type() for m in payload]
            return [{
                'name': msg.get_filename(),
                'content': payload_get_type( payload, types, ['text/html', 'text/plain'] ),
                'type': 'text',
            }]
        else:
            return list( flatten( [parse_email( p ) for p in payload] ) )
    elif ( content_type == 'message/rfc822' ):
        return list( flatten( [parse_email( p ) for p in payload] ) )
    elif ( content_type == 'application/pgp-signature' ):
        return [{
            'name': None, # Include in body
            'content': msg.get_payload().replace( '\n', '<br>' ),
            'type': 'text',
        }]
    elif ( content_type != 'message/delivery-status' ):
        return [{
            'name': msg.get_filename(),
            'content': msg.get_payload( decode=True ),
            'type': content_type,
        }]

def filler_message( mid ):
    m = mailbox.Message()
    m['message-id'] = mid
    m['subject'] = '[Not in archive]'
    m['date'] = '[Not in archive]'
    m['from'] = ''
    return m

def safely_append_thread( mid, par, threads, messages ):
    if ( par is None ): return
    if ( mid not in threads ):
        threads[mid] = []
    if ( par not in threads ):
        threads[par] = []
    if ( mid not in threads[par] ):
        threads[par].append( mid )
    # Creates filler if missing
    if ( par not in messages ):
        messages[par] = filler_message( par )
    if ( mid not in messages ):
        messages[mid] = filler_message( mid )
    if ( messages[mid].get( 'in-reply-to' ) is None ):
        messages[mid]['in-reply-to'] = par

def get_parent_id( msg ):
    irt = msg.get( 'in-reply-to' )
    if ( irt is not None ): return irt
    refs = msg.get( 'references' )
    if ( refs is not None ): return refs[-1]
    return None

# Establishes hierarchical thread relations
# (Also modifies in in-reply-to field if needed for later use)
# TODO: Implement https://www.jwz.org/doc/threading.html
# Current implementation assumes Message-IDs are consistent between parent and
# child (which is not always the case), and requires manual intervention in
# this case, and in the case where fields aren't filled out consistently
def get_threads( messages ):
    threads = {} # Contains direct children
    for mid, msg in messages.copy().items():
        if ( mid not in threads ):
            threads[mid] = []
        # Adds IRT content (if available)
        irt = msg.get( 'in-reply-to' )
        safely_append_thread( mid, irt, threads, messages )
        # Adds references content (if available)
        # References are (typically) hierarchical: 1st is parent of 2nd, 2nd of 3rd, etc.
        # TODO: Not guaranteed to be separated by spaces
        refs = re.findall( r'\S+', msg.get( 'references' ) or '' )
        if ( ( irt is not None ) and ( irt not in refs ) ): refs.append( irt )
        for parent, child in zip( refs, refs[1:] ):
            safely_append_thread( child, parent, threads, messages )

    # Pruning pass - ensure each thread only has 1 parent
    for mid, children in threads.items():
        for c in children.copy():
            if ( get_parent_id( messages[c] ) != mid ):
                children.remove( c )

    # TODO: Find dead roots; attempt to connect to other threads

    return threads

def content_to_html( msg, content, threads, messages, outdir, body_path ):
    if ( content is None ): return

    # Writes body of html/header info
    with open( body_path, 'w' ) as file:
        file.write( '''
            <html>
                <head>
                    <title>%s</title>
                </head>
                <body>
                    <p><a href="index.html">Index</a></p>
                    <p><strong>Subject</strong>: %s</p>
                    <p><strong>From</strong>: %s</p>
                    <p><strong>Date</strong>: %s</p>
        ''' % (
            '%s - %s - %s' % (
                html.escape( get_header_text( msg, 'subject' ) ),
                html.escape( format_date( msg ) ),
                'Email Archive'
            ),
            html.escape( get_header_text( msg, 'subject' ) ),
            html.escape( get_header_text( msg, 'from' ) ),
            html.escape( format_date( msg ) ),
        ) )

        # Parent info
        parent = get_parent_id( msg )
        if ( parent is not None ):
            if ( parent in messages ):
                file.write( '''
                        <p><a href="%s">Parent</a></p>
                ''' % ( urllib.parse.quote( parent ) + '.html' ) )
            else:
                file.write( '''
                        <p><em>Parent not archived</em></p>
                ''' )

    # Writes message content
    attachments = []
    for part in content:
        if ( part is None ):
            continue
        name = part['name']
        # Append for multi-part messages/body
        if ( name is None ):
            name = msg_id + '.html'
            filepath = body_path
            # hr to distinguish content
            part['content'] = '<hr>' + part['content']
        # For attachments, just create the director if needed
        else:
            os.makedirs( attachment_path, exist_ok=True )
            filepath = os.path.join( attachment_path, name )
            attachments.append( { 'name': name, 'path': filepath } )

        if part['type'] == 'text':
            # TODO: Attempt to detect encoding?
            try:
                open( filepath, 'a' ).write( part['content'] )
            except UnicodeEncodeError:
                open( filepath, 'a', encoding='utf8' ).write( part['content'] )
        else:
            open( filepath, 'ab' ).write( part['content'] )

    # Finishes body/writes footer info
    with open( body_path, 'a' ) as file:
        if ( len( attachments ) > 0 ):
            # Attachments portion of footer
            file.write( '''
                    <hr>
                    <p><strong>Attachments</strong>:</p>
                    <p><em>(Please be wary of attachments - they have not been scanned for viruses)</em></p>
                    <ul>
                        %s
                    </ul>
            ''' % (
                '\n'.join(
                    [
                        '<li><a href="../%s">%s</a>' % ( urllib.parse.quote( a['path'] ), a['name'] )
                        for a in attachments
                    ]
                )
            ) )

        # Replies
        if ( len( threads[msg_id] ) > 0 ):
            file.write( '''
                    <hr>
                    <p><strong>Replies</strong>:</p>
                    <ul>
                        %s
                    </ul>
            ''' % (
                '\n'.join(
                    [
                        '<li><a href="%s">%s</a>' % (
                            urllib.parse.quote( child + '.html' ),
                            html.escape( get_header_text( messages[child], 'subject' ) ),
                        )
                        for child in threads[msg_id]
                    ]
                )
            ) )

        # Finishes body
        file.write( '''
                </body>
            </html>
        ''' )

def write_message_tree( file, msg_ids, threads, messages ):
    for mid in msg_ids:
        msg = messages[mid]
        file.write( '<li>%s: %s</li>' % (
            html.escape( format_date( msg ) ),
            '<a href="%s">%s</a>' % (
                urllib.parse.quote( mid + '.html' ),
                get_header_text( msg, 'subject' )
            )
        ) )
        if ( len( threads[mid] ) > 0 ):
            file.write( '<ul>' )
            write_message_tree( file, threads[mid], threads, messages )
            file.write( '</ul>' )

def sort_helper( msg, messages ):
    if ( isinstance( msg, str ) ):
        return sort_helper( messages[msg], messages )
    # 10 = number of elements in date tuple
    # TODO: Potentially infer time?
    return email.utils.parsedate_tz( msg.get( 'date' ) ) or 10 * (math.inf,)

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('mboxfile', default='export.mbox', help='path to mbox file')
    parser.add_argument('outdir', default='email-archive', help='output directory')
    parser.add_argument('-f', '--recipient-filter', help='filter by to/cc:/rto: recipient (e.g. by mailinglist)')
    parser.add_argument('-m', '--mode', choices=['plain','html'], default='html', help='multipart mime-type to extract, default: text/html')
    args = parser.parse_args()

    filename = args.mboxfile
    outdir = args.outdir
    recipient_filter = args.recipient_filter
    mode = args.mode

    if not os.path.isfile(filename):
        print(f'path "{filename}" is not a file')
        exit(1)

    if not os.path.exists(outdir):
        os.makedirs(outdir)

    mbox = mailbox.mbox( filename )
    messages = {}
    for key, msg in mbox.items():
        to = msg.get( 'to' ) or msg.get( 'delivered-to' ) or ''
        cc = msg.get( 'cc' ) or ''
        rto = msg.get( 'reply-to' ) or ''
        if ( recipient_filter ):
            if ( ( to.find( recipient_filter ) >= 0 )
              or ( cc.find( recipient_filter ) >= 0 )
              or ( rto.find( recipient_filter ) >= 0 ) ):
                messages[msg.get( 'message-id' )] = msg
        elif msg.get('message-id'):
            messages[msg.get( 'message-id' )] = msg
        else:
            print(f'"{key}" has no message-id')
            continue

    # Gets parental info
    threads = get_threads( messages )

    # Sorts replies to be in order
    for mid, msgs in threads.copy().items():
        threads[mid].sort( key = lambda x: sort_helper( x, messages ) )

    # Writes email html files
    for key, msg in messages.items():
        msg_id = msg.get( 'message-id' )
        body_path = os.path.join( outdir, msg_id + '.html' )
        attachment_path = os.path.join( outdir, msg_id )

        # Deletes all previous files (if they exist) for easier append-age later
        try:
            os.remove( body_path )
        except OSError:
            pass
        shutil.rmtree( attachment_path, ignore_errors=True )

        content = parse_email( msg )
        content_to_html( msg, content, threads, messages, outdir, body_path )

    # Writes index.html
    # Sorts files based on timestamp; makes things easier
    sorted_messages = [x for x in messages.values()]
    sorted_messages.sort( key = lambda x: sort_helper( x, messages ) )

    roots = [
        f.get( 'message-id' ) for f in sorted_messages
            if get_parent_id( f ) is None
    ]

    with open( os.path.join( outdir, 'index.html' ), 'w' ) as file:
        file.write( '''
        <html>
            <head>
                <title>Email Archive</title>
            </head>
            <body>
                <ul>''' )

        write_message_tree( file, roots, threads, messages )

        file.write( '''
                </ul>
            </body>
        </html>
        ''' )
