#!/usr/bin/env python
"""
Given a directory path, this searches all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""

import argparse
import docx2txt
import fnmatch
import os

stats = { 'matched': 0,
         'searched': 0 }

# Your awesome code begins here!
def scan_for_dotms( dir = '.' ):
    return filter( lambda n: fnmatch.fnmatch( n, '*.dotm' ) , map( lambda w: dir + '/' + w , list( os.listdir( dir ) ) ) )

def process_dotms( dotms ):
    return map( lambda d: ( d, docx2txt.process( d ) ), dotms )

def search_dotms( dotms, text ):
    global stats
    stats[ 'searched' ] = len( dotms )
    return filter( lambda u: u[2] > -1, map( lambda w: ( w[0], w[1], w[1].find( text ) ), dotms ) )

def print_dotms( dotms ):
    global stats
    stats[ 'matched' ] = len( dotms )
    for dotm in dotms:
        if dotm[2] < 40:
            mark = 0
        else:
            mark = dotm[2] - 40
        print( dotm[0] + '\n\n\t' + dotm[1][ mark : dotm[2] + 40 ] + '\n\n' )
    print( 'Found {m:d} matches in {s:d} files'.format( m=stats['matched'], s=stats['searched'] ) )

def main():
    parser = argparse.ArgumentParser( description='Given a directory path, this searches all files in the path for a given text string within the \'word/document.xml\' section of a MSWord .dotm file.' )
    parser.add_argument( '--dir', default='.', type=str )
    parser.add_argument( 'string', type=str )
    args = parser.parse_args()

    if args.dir:
        print_dotms( search_dotms( process_dotms( scan_for_dotms( args.dir ) ), args.string ) )
    else:
        print_dotms( search_dotms( process_dotms( scan_for_dotms( '.' ) ), args.string ) )


if __name__ == '__main__':
    main()