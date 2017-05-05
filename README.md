# easyconvert
Excel to CSV conversion Tool

Libreoffice/Openoffice provides graphical utility to convert excel files to csv and vice versa.  But its not practical while dealing with bulk list of files. In such cases, it will be better if we can handle the same via command-line.

Easycovert is a program written in Python which helps to do the conversions thorough single shell command. Its a faster utility and uses a better method in separating the sheets from excel files. This program requires Python3 environment for the proper working.

#Install Steps

Step1: Download easyconvert script

Step2: Copy it to Linux binary location

#Usage

usage: easyconvert [-h] [-csv] [-xls] [-xlsx] filename

Convert xls/xlsx to csv and vice versa

positional arguments:
filename    Filename

optional arguments:
-h, –help  show this help message and exit
-csv        Convert xls/xlsx to csv
-xls        Convert csv to xls
-xlsx       Convert csv to xlsx
