import mailbox
import email, email.policy, email.utils
import html
import chardet
import base64
import struct
import copy
import re
import os
import shutil

def flatten( l ):
    for i in l:
        if isinstance( i, list ):
            yield from flatten( i )
        else:
            yield i

# Puts date into consistent format
def format_date( msg ):
    date = msg.get( 'date' )
    return email.utils.format_datetime( email.utils.parsedate_to_datetime( date ) )

# Helps extracting tricky header info (so far only needed for subjects)
def get_header_text( msg, item ):
    i = msg.get( item )
    if ( isinstance( i, str ) ):
        return i
    elif ( isinstance( i, email.header.Header ) ):
        sub = email.header.decode_header( i )[0][0] # TODO: What about others?
        return sub.decode( chardet.detect( sub )['encoding'] )
    else:
        print( 'TODO' )

def get_payload_text( msg ):
    subtype = msg.get_content_subtype()
    payload = msg.get_payload( decode=True )
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

def get_parent_id( msg ):
    references = msg.get( 'references' )
    if ( references is None ):
        # Try in-reply-to
        irt = msg.get( 'in-reply-to' )
        if ( irt is not None ): return irt
        # # TODO: Try thread-index
        # # https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxomsg/9e994fbb-b839-495f-84e3-2c8c02c7dd9b
        # thread_index = msg.get( 'thread-index' )
        # if ( thread_index is None ): return None
        # # Extracts thread index data
        # thread_index = base64.b64decode( thread_index )
        # filetime = struct.unpack( '>xIB', thread_index[:6] )
        # # I have some confusion here, since the docs say the GUID has data 1-3,
        # # but it's 16 bytes, so how is it divided? Searching yielded no further
        # # confusion, as the only other "p"-guid I could find is the packet
        # # guid, which has 4 data fields
        # # https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-oleps/5ee5aa9d-6b96-4e54-a6bc-6b1f562d616b
        # guid = struct.unpack( '>IHHQ', thread_index[6:22] )
        # response_levels = []
        # for r in range( 22, len( thread_index ), 5 ):
        #     response_levels.append( struct.unpack( '>I', thread_index[r:r + 4] )[0] )
        # # TODO: or just do https://www.jwz.org/doc/threading.html
        return None
    # Assumes last references is the replied-to email
    return re.split( r'\s+', references )[-1]

def content_to_html( msg, content, children, message_ids, outdir ):
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
                'LUG @ NC State Email Archive'
            ),
            html.escape( get_header_text( msg, 'subject' ) ),
            html.escape( msg.get( 'from' ) ),
            html.escape( format_date( msg ) ),
        ) )

        # Parent info
        parent = get_parent_id( msg )
        if ( parent is not None ):
            if ( parent in message_ids ):
                file.write( '''
                        <p><a href="%s">Parent</a></p>
                ''' % ( parent + '.html' ) )
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
        with open( filepath, 'a' if part['type'] == 'text' else 'ab' ) as file:
            file.write( part['content'] )

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
                        '<li><a href="../%s">%s</a>' % ( a['path'], a['name'] )
                        for a in attachments
                    ]
                )
            ) )

        # Replies
        if ( len( children[msg_id] ) > 0 ):
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
                            child + '.html',
                            html.escape( get_header_text( message_ids[child], 'subject' ) ),
                        )
                        for child in children[msg_id]
                    ]
                )
            ) )

        # Finishes body
        file.write( '''
                </body>
            </html>
        ''' )

def write_message_tree( file, msg_ids, children, message_ids ):
    for msg_id in msg_ids:
        msg = message_ids[msg_id]
        file.write( '<li>%s: %s</li>' % (
            html.escape( format_date( msg ) ),
            '<a href="%s">%s</a>' % (
                msg_id + '.html',
                get_header_text( msg, 'subject' )
            )
        ) )
        if ( len( children[msg_id] ) > 0 ):
            file.write( '<ul>' )
            write_message_tree( file, children[msg_id], children, message_ids )
            file.write( '</ul>' )

if __name__ == '__main__':
    filename = 'export.mbox'
    outdir = 'out'

    # TODO: Check if file exists
    mbox = mailbox.mbox( filename )
    messages = {}
    for key, msg in mbox.items():
        to = msg.get( 'to' ) or msg.get( 'delivered-to' )
        if ( to.find( 'lug@lists.ncsu.edu' ) >= 0 ):
            messages[key] = msg

    # Gets parental info
    children = {}
    message_ids = {}
    for key, msg in messages.items():
        msg_id = msg.get( 'message-id' )
        message_ids[msg_id] = msg
        if ( msg_id not in children ):
            children[msg_id] = []
        parent = get_parent_id( msg )
        if ( parent not in children ):
            children[parent] = []
        if msg_id not in children[parent]:
            children[parent].append( msg_id )

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
        content_to_html( msg, content, children, message_ids, outdir )

    # Writes index.html
    # 1. Sort files based on timestamp (by message_id to eliminate dupes)
    sorted_messages = [x for x in message_ids.values()]
    sorted_messages.sort( key = lambda m: email.utils.parsedate_tz( m.get( 'date' ) ) )
    # 2. Find messages that either have no parent or parent html file
    roots = [
        f.get( 'message-id' ) for f in sorted_messages if not get_parent_id( f ) \
            or not os.path.exists( os.path.join(
                outdir, get_parent_id( f ) + '.html'
            ) )
    ]
    with open( os.path.join( outdir, 'index.html' ), 'w' ) as file:
        file.write( '''
        <html>
            <head>
                <title>LUG @ NC State Email Archive</title>
            </head>
            <body>
                <ul>''' )

        write_message_tree( file, roots, children, message_ids )

        file.write( '''
                </ul>
            </body>
        </html>
        ''' )
