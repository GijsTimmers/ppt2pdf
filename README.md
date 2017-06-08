# ppt2pdf
A wrapper to convert MS Office `ppt` or `pptx` files to `pdf` files

## Purpose
This script is intended to help students using GNU Linux.

Sadly enough, professors often distribute their class presentations
in the closed-source `ppt` or `pptx` format. These can be opened with 
LibreOffice nowadays, but it's not quite the same: the fonts often don't match
and figures are misaligned.

This script will end your misery by batch-converting the presentations in
the pdf format, which albeit closed-source too renders much better on Linux.

## Dependencies
- Microsoft Windows 7 or higher
- Microsoft Office 2013 or higher
- Python > 2.5
- pypiwin32: `sudo pip install pypiwin23`

## Usage
Navigate your shell into a directory containing the `ppt` or pptx` files.
Then, run:

     ./ppt2pdf
 
A directory called `pdf` will be created if it doesn't exist yet, containing
the converted files. 

## Disclaimer
This script is a wrapper for the code in 
[K. Scott Allen's](http://odetocode.com/about/scott-allen) 
[blog post](http://odetocode.com/blogs/scott/archive/2013/06/26/convert-a-directory-of-powerpoint-slides-to-pdf-with-python.aspx)

