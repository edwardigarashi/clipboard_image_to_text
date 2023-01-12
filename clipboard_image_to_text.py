#!/usr/bin/env python   
import ctypes,sys,os,traceback, optparse,re,time
from PIL import ImageGrab,Image
from pytesseract import pytesseract
"""
SYNOPSIS
    TODO clipboard_image_to_text.py [-h,--help] [-v,--verbose] [--version]

DESCRIPTION
    TODO Copy an image to your clipboard on windows and run the script or
    if the user requests help (-h or --help).

EXAMPLES
    TODO: python.exe clipboard_image_to_text.py
    >>> This is a test message

EXIT STATUS
    TODO: List exit codes

AUTHOR
    TODO: Edward Igarashi <info@igarashi.net>

LICENSE
    MIT License

    Copyright (c) 2023 Edward Igarashi

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.

VERSION
    v1.0
"""
def set_clipBoard(text):
    command = 'echo | set /p nul=' + text.strip() + '| clip'
    os.system(command)

def read_image():
    try:
        path_to_tesseract = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
        img = ImageGrab.grabclipboard()
        pytesseract.tesseract_cmd = path_to_tesseract
        text = pytesseract.image_to_string(img)
        return text
    except TypeError as e:
        print(e)
        main() if ctypes.windll.user32.MessageBoxW(0, "Please copy a image inside the clipboard and click retry", "Item in clipboard not image!", 5) == 4 else print("Canceled")
        sys.exit(0)
    except Exception as e:
        print('ERROR, UNEXPECTED EXCEPTION')
        print(str(e))
        traceback.print_exc()
        os._exit(1)    
        
def main ():
    try:
        global options, args
        response = 10
        while response == 10 or response == 11:
            text = read_image()
            title = "Clipboard Image to Text"
            if response == 11:
                set_clipBoard(text)
                title = "Copied to Clipboard!"
            response = ctypes.windll.user32.MessageBoxW(0, text, title, 6)
    except Exception as e:
        print('ERROR, UNEXPECTED EXCEPTION')
        print(str(e))
        traceback.print_exc()
        os._exit(1)    

if __name__ == '__main__':
    try:
        start_time = time.time()
        parser = optparse.OptionParser(formatter=optparse.TitledHelpFormatter(), usage=globals()['__doc__'], version='$Id$')
        parser.add_option ('-v', '--verbose', action='store_true', default=False, help='verbose output')
        (options, args) = parser.parse_args()
        #if len(args) < 1:
        #    parser.error ('missing argument')
        if options.verbose: print(time.asctime())
        main()
        if options.verbose: print(time.asctime())
        if options.verbose: print('TOTAL TIME IN MINUTES:',)
        if options.verbose: print((time.time() - start_time) / 60.0)
        sys.exit(0)
    except KeyboardInterrupt as e: # Ctrl-C
        raise e
    except SystemExit as e: # sys.exit()
        raise e
    except Exception as e:
        print('ERROR, UNEXPECTED EXCEPTION')
        print(str(e))
        traceback.print_exc()
        os._exit(1)