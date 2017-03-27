# ConvertToXLSX

A python script that converts an image file to an excel spreadsheet by changing the background colour of the spreadsheet cells. It can also change the colour of the image from 24-bit to 4-bit.

It should be noted that I have only tried this on small images (i.e. less than 550x550 pixels in size) changing the background of so many cells makes the excel file slow to load and use, so be patient. The image is best viewed in excel at 10% Zoom.

Why create this? Just for fun, it is pretty useless, however some possible uses include teaching people about how images are stored as pixels and how those pixels can be accessed as a 2D-array. Or it could be used to create art.

## Installation

Python 2.7.10

XlsxWriter and Python Image Library (PIL) are required - both can be installed using pip

## Usage

$ python convert.py "image file location"

OR

$ python convert.py "image file location" -d

The -d argument changes the colour from 24-bit to 4-bit.

## Credits

Created by: Joel Sprinkle

## License

This work is licensed under a [Creative Commons Attribution-ShareAlike 3.0 Unported License](http://creativecommons.org/licenses/by-sa/3.0/).
