#!/usr/bin/python2.7

import os
import argparse
import math
import xlsxwriter
from PIL import Image


# ---- functions
# -- Checks to make sure the image file passed in is a real file of the appropriate type
def check_img_file(img_path):
    if not os.path.isfile(img_path):
        print('Valid file path not given')
        exit(1)

    ext = os.path.splitext(img_path)[1].lower()  # valid extensions: BMP, JPG, JP2, PNG, TIFF
    if ext != '.bmp' and ext != '.jpg' and ext != '.jp2' and ext != '.png' and ext != '.tiff':
        print('"' + ext + '" is not a valid file type, only bmp, jpg, jp2, png, or tiff file types are allowed')
        exit(1)


# -- Convert colours to 16 base colours
def convert_colours(pixel):
    smallest_dif = 10000
    new_colour = ''
    # list of available colours from: http://xlsxwriter.readthedocs.io/working_with_colors.html#colors
    four_bit = [[[0, 0, 0], 'black'],
                [[0, 0, 255], 'blue'],
                [[128, 0, 0], 'brown'],
                [[0, 255, 255], 'cyan'],
                [[128, 128, 128], 'gray'],
                [[0, 128, 0], 'green'],
                [[0, 255, 0], 'lime'],
                [[255, 0, 255], 'magenta'],
                [[0, 0, 128], 'navy'],
                [[255, 102, 0], 'orange'],
                [[255, 0, 255], 'pink'],
                [[128, 0, 128], 'purple'],
                [[255, 0, 0], 'red'],
                [[192, 192, 192], 'silver'],
                [[255, 255, 0], 'yellow'],
                [[255, 255, 255], 'white']]

    for i in range(0, len(four_bit)):
        # https://en.wikipedia.org/wiki/Color_difference
        d = math.sqrt(((pixel[0] - four_bit[i][0][0]) ** 2) +
                      ((pixel[1] - four_bit[i][0][1]) ** 2) +
                      ((pixel[2] - four_bit[i][0][2]) ** 2))
        if d < smallest_dif:
            smallest_dif = d
            new_colour = four_bit[i][1]
    return new_colour


# ---- starting point
# -- Command line arguments
parser = argparse.ArgumentParser(description='Converts an image to an xlsx file.', prog='ConvertToXLSX')
parser.add_argument('-v', '--version', action='version', version='%(prog)s 1.0')
parser.add_argument('imgPath', help='the file path to the image to be converted to an xlsx file')
parser.add_argument('-d', '--decolourize', action='store_true', default=False,
                    help='converts the image from 24-bit to 4-bit colours')

args = parser.parse_args()
imgPath = args.imgPath
decolourize = args.decolourize

check_img_file(imgPath)  # check valid image file was passed in

# -- Load the image
# http://stackoverflow.com/questions/138250/how-can-i-read-the-rgb-value-of-a-given-pixel-in-python
im = Image.open(imgPath)
pix = im.load()

# -- Load excel workbook
# http://xlsxwriter.readthedocs.io/index.html
workbook = xlsxwriter.Workbook(os.path.splitext(imgPath)[0] + '.xlsx')  # use the images base name
worksheet = workbook.add_worksheet()
worksheet.set_column(0, im.size[0], 0.5)  # make the cells more pixel shaped

# -- Convert the image
print('Converting image to excel document')
for column in range(0, im.size[0]):
    for row in range(0, im.size[1]):
        if column % 500 == 0 and row % 500 == 0:  # Just to show that the program is still working
            print('...')

        if decolourize:
            formatted_colour = convert_colours(pix[column, row])
            if formatted_colour == '':
                print('Error transforming colour from 24-bit to 4-bit')
                exit(1)
        else:
            formatted_colour = '#%02x%02x%02x' % pix[column, row]  # convert RGB to hex value - does not check bounds

        colour = workbook.add_format({'bg_color': formatted_colour})
        worksheet.write(row, column, '', colour)

print('Finishing up')
workbook.close()
im.close()
