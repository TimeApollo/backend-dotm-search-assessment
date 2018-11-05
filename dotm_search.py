#!/usr/bin/env python
"""
Given a directory path, this searches all files in the path for a given text string 
within the 'word/document.xml' section of a MSWord .dotm file.
"""

# Your awesome code begins here!

import zipfile
import os
import argparse

def dir_text_search( directory , search_text ):
    """looks for .dotm files in directory and then searches them for input text"""

    print("Searching directory {} for text {}").format(directory , search_text)
    # document = zipfile.ZipFile('dotm_files/G003.dotm').namelist()

    searched_files = 0
    matched_files = 0
    for file in os.listdir( directory ):
        if file.endswith('.dotm'):
            searched_files += 1
            full_path = os.path.join(directory , file)
            document = zipfile.ZipFile(full_path)
            xml_content = document.read('word/document.xml')
            document.close()
            
            index = xml_content.find(search_text)

            if index >= 0:
                matched_files += 1
                print("Match found in file {}").format(full_path)
                print('...' + xml_content[index - 40: index + 40] + '...')

    print('Total dotm files searched: {}').format(searched_files)
    print('Total dotm files matched: {}').format(matched_files)


def arg_parser():
    """takes in 1 required text arg and 1 optional dir arg."""
    """"""
    """uses current dir if optional dir arg not supplied"""
    
    # initiate argparse
    parser = argparse.ArgumentParser()
    # add arguments
    parser.add_argument('-d' , '--dir' , type=str, default='./', help='set directory to search')
    parser.add_argument("text" , type=str , help='text to search for in documents')
    # return parser to be used
    args = parser.parse_args()
    return args

def main():
    """sets up arg parser. calls text search function"""
    args = arg_parser()

    if os.path.isdir(args.dir):
        dir_text_search(args.dir , args.text)
    else:
        print('{} is not a directory').format(args.dir)
        print('Please provide a directory to be searched.')


if __name__ == '__main__':
    main()
