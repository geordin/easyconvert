# easyconvert
Excel to CSV conversion Tool

Libreoffice/Openoffice provides graphical utility to convert excel files to csv and vice versa.  But its not practical while dealing with bulk list of files. In such cases, it will be better if we can handle the same via command-line.

Easycovert is a program written in Python which helps to do the conversions thorough single shell command. Its a faster utility and uses a better method in separating the sheets from excel files. Since the Easyconvert program contains all the required Python libraries, it will be compatible in all Linux environments with out any additional configurations.

Install Steps

Step1: Download the package

#wget http://techies-world.com/wp-content/uploads/2017/05/easyconvert.tar.gz

Step2: Extract the package

#tar -xvf easyconvert.tar.gz

Step3: Change the location to the extracted folder

#cd easyconvert

Step4: Run the install script

#sh install.sh

Uninstall Steps

Step1: Download the package

#wget http://techies-world.com/wp-content/uploads/2017/05/easyconvert.tar.gz

Step2: Extract the package

#tar -xvf easyconvert.tar.gz

Step3: Change the location to the extracted folder

#cd easyconvert

Step4: Run the uninstall script

#sh uninstall.sh

Usage

usage: easyconvert [-h] [-csv] [-xls] [-xlsx] filename

Convert xls/xlsx to csv and vice versa

positional arguments:
filename    Filename

optional arguments:
-h, –help  show this help message and exit
-csv        Convert xls/xlsx to csv
-xls        Convert csv to xls
-xlsx       Convert csv to xlsx
