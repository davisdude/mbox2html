import bs4
import os
import re

# Removes area codes from phone numbers and TLDs from emails
def replace( string ):
    string = string.replace( '<br>', '\n' )
    # Removes area code from phone numbers
    # Only handles US-style (+country-(area)-(xxx)-(xxxx)) or plain-style (no
    # separators)
    # TODO: Exclude URLs
    phone_find = re.compile( r'''
        (\s+|^)                             # Don't match in the middle (e.g. URL)
                                (\+\s*\d+)? # Country code
                  (\s|\-|\.|/)*             # Separator
        (\()?(\s*)                          # Area code open paren
                                (\d{3})     # Area code
        (\s*)(\))?                          # Area code close paren
                  (\s|\-|\.|/)*             # Separator
                                (\d{3})     # Next 3 digits
                  (\s|\-|\.|/)*             # Separator
                                (\d{4})     # Final 4 digits
    ''', re.X )
    phone_repl = r'\1\2\3\4\5XXX\7\8\9\10\11\12'
    string = phone_find.sub( phone_repl, string )

    # Removes email addresses
    email_find = re.compile( r'''
        ([a-zA-Z0-9_.+-]+@[a-zA-Z0-9-.]+)(\.[a-zA-Z0-9-]+)
    ''', re.X )
    email_repl = r'\1.[redacted]'
    string = email_find.sub( email_repl, string )

    return string

def recursive_replace( soup ):
    if ( type( soup ) == str ):
        return replace( soup )
    elif ( not hasattr( soup, 'contents' ) ):
        return soup.__class__( replace( soup.string ) )
    else:
        for c in soup.contents:
            c.replace_with( recursive_replace( c ) )
        for k, v in soup.attrs.items():
            # Only worry about mailto (for now, at least)
            if ( k != 'href' ): continue
            if ( re.match( r'^mailto:', v, re.I ) is None ): continue
            soup[k] = recursive_replace( v )
    return soup

if __name__ == '__main__':
    for entry in os.scandir( 'out' ):
        if ( entry.name.endswith( '.html' ) and entry.is_file() ):
            with open( entry.path, 'r' ) as file:
                soup = bs4.BeautifulSoup( file, 'html.parser' )
            recursive_replace( soup )
            with open( entry.path, 'w' ) as file:
                file.write( soup.prettify() )
        elif ( entry.is_dir() ):
            for att in os.scandir( entry.path ):
                if ( att.name.endswith( '.vcf' ) and att.is_file() ):
                    with open( att.path, 'r' ) as file:
                        text = file.read()
                    text = replace( text )
                    with open( att.path, 'w' ) as file:
                        file.write( text )
