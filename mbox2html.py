# TODO: Header file/templates
#       Checking of input/output files/directories
#       Email attachments

import mailbox
import html
import dateutil.parser as date_parser
import argparse
import os
import re

title = 'LUG @ NC State Email Archives'

# Removes TLD from emails for privacy
pattern = re.compile(r'([\w\.\+_-]+@[\w\._-]+\.)([a-zA-Z]+)')
def filter_emails(text):
    return pattern.sub(
        r'\1&lt;hidden&gt;',
        text
    )

def get_content(message, type):
    for part in message.walk():
        if part.get_content_type() == type:
            content = part.get_payload(decode=True)
            charset = part.get_content_charset()
            print(charset)
            return content.decode(charset or 'utf-8')
    return None

def get_body(message):
    types = ['text/html', 'text/plain']
    for t in types:
        body = get_content(message, type=t)
        if body is not None:
            break
    return body

def read_email(message, emails):
    emails[message['Message-ID']] = message

def populate_responses(message, responses):
    irt = message.get('In-Reply-To')
    if irt is None:
        return
    if irt not in responses:
        responses[irt] = []
    if message['Message-ID'] not in responses[irt]:
        responses[irt].append(message['Message-ID'])

def html_helper(message, key=None, escape=True, emails=True):
    if key is not None:
        message = message.get(key)
        if escape:
            message = html.escape(message)
        if emails:
            message = filter_emails(message)
        return '<p><b>%s</b>: %s</p>\n' % (key, message)
    else:
        if escape:
            message = html.escape(message)
        if emails:
            message = filter_emails(message)
        return '<p>%s</p>\n' % (message)

def link(url, text):
    return '<a href="%s">%s</a>' % (html.escape(url), text)

def tree(e):
    mid = e['Message-ID']
    date = date_parser.parse(e['Date']).strftime('%B %d, %Y')
    s = '<li>%s: %s</li>' % (date, link(mid + '.html', e['Subject']))
    if mid in responses:
        s += '<ul>\n'
        for r in responses[mid]:
            s += tree(emails[r])
        s += '</ul>\n'
    return s

def handle_ascii_quotes(text):
    quotes = []
    any_quotes = False
    in_quote = None
    for line in text.split('\r\n'):
        if line.startswith('>'):
            if not in_quote:
                any_quotes = True
                in_quote = True
                quotes.append([])
        else:
            in_quote = False
        if in_quote:
            quotes[len(quotes) - 1].append(line[1:])
        else:
            quotes.append(line)

    new_text = ''
    for l in quotes:
        if isinstance(l, str):
            new_text += '%s\r\n' % filter_emails(l)
        else:
            print(l)
            new_text += '<blockquote>%s</blockquote>' % filter_emails(
                '\r\n'.join(l)
            )
    if any_quotes:
        return handle_ascii_quotes(new_text)
    else:
        return new_text

###############################################################################

long_description = 'Generates a collection of HTML files from a given mbox ' \
                 + 'file in the style of an online email archive'

if __name__ == '__main__':
    # Parses args
    parser = argparse.ArgumentParser(description=long_description)
    parser.add_argument(
        '-i', '--infile',
        type=str,
        help='Input mbox file'
    )
    parser.add_argument(
        '-o', '--outdir',
        type=str,
        help='Output directory for HTML files'
    )
    parser.add_argument(
        '-l', '--list',
        type=str,
        help='Intended for use with mailing lists; use this option to ' \
           + 'exclude emails not explicitly sent to the given email address',
        nargs='?',
        default=''
    )
    args = parser.parse_args()

    # Associates responses with emails
    emails = {}
    responses = {}
    mb = mailbox.mbox(args.infile)

    # Sorts emails by date
    mb = [x for x in mb]
    mb.sort(key = lambda message: date_parser.parse(message['Date']))

    for message in mb:
        # Filters to only emails sent to mailing list
        if message['To'].find(args.list) < 0:
            continue
        read_email(message, emails)
        populate_responses(message, responses)

    # Creates index.html
    with open(os.path.join(args.outdir, 'index.html'), 'w') as f:
        f.write('<title>%s</title>' % title)
        f.write('<ul>\n')
        parent_threads = [e for e in emails.values() if not e['In-Reply-To']]
        for t in parent_threads:
            f.write(tree(t))
        f.write('</ul>\n')

    # Creates HTML files
    for mid, message in emails.items():
        print(mid)
        with open(os.path.join(args.outdir, mid + '.html'), 'w') as f:
            date, subject = message.get('Date'), message.get('Subject')
            f.write('<title>%s - %s - %s</title>' % (subject, date, title))
            f.write(html_helper(link('index.html', 'Index'), escape=False))
            f.write(html_helper(message, 'Subject'))
            f.write(html_helper(message, 'From'))
            f.write(html_helper(message, 'Date'))

            if message['In-Reply-To'] is not None:
                f.write(html_helper(link(message['In-Reply-To'] + '.html', 'Parent'), escape=False, emails=False))

            f.write(html_helper('<hr>', escape=False))
            # Handles ASCII replies better
            body = get_body(message)
            for p in handle_ascii_quotes(body).split('\r\n\r\n'):
                f.write('<p>%s</p>' % filter_emails(p))
            f.write(html_helper('<hr>', escape=False))

            if mid in responses:
                f.write(html_helper('<b>Replies:</b>', escape=False))
                f.write('<ul>\n')
                for r in responses[mid]:
                    f.write('<li>%s</li>\n' % link(emails[r]['Message-ID'] + '.html', emails[r]['Subject']))
                f.write('</ul>\n')
