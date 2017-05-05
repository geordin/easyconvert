#!/usr/bin/python3

import xlrd, csv, os, argparse, glob, datetime
from xlsxwriter.workbook import Workbook

parser = argparse.ArgumentParser(description="Convert xls/xlsx to csv and vice versa")
parser.add_argument("-csv", action='store_true', help="Convert xls/xlsx to csv")
parser.add_argument("-xls", action='store_true', help="Convert csv to xls")
parser.add_argument("-xlsx", action='store_true', help="Convert csv to xlsx")
parser.add_argument("filename", type=str, help="Filename")
args = parser.parse_args()

filename = args.filename
name = filename.split('.')[0]

def xlstocsv():
	workbook = xlrd.open_workbook(filename)
	sheet_no=workbook.nsheets
	count=0
	while (count < sheet_no):
		count += 1
		worksheet = workbook.sheet_by_index(0)
		num_rows = worksheet.nrows - 1
		num_cells = worksheet.ncols - 1
		curr_row = -1
		while curr_row < num_rows:
			curr_row += 1
			row = []
			curr_cell = -1
			while curr_cell < num_cells:
				curr_cell += 1
				# Cell Types: 0=Empty, 1=Text, 2=Number, 3=Date, 4=Boolean, 5=Error, 6=Blank
				cell_type = worksheet.cell_type(curr_row, curr_cell)
				if cell_type == 3:
					t = xlrd.xldate_as_tuple((worksheet.cell_value(curr_row, curr_cell)), workbook.datemode)
					if t[0] == t[1] == t[2] == 0:
						cell_value = datetime.time(t[3], t[4], t[5])
					elif t[3] == t[4] == t[5] == 0:
						cell_value = datetime.date(t[0], t[1], t[2])
					else:
						cell_value = str(datetime.datetime(t[0], t[1], t[2], t[3], t[4], t[5]))
				elif cell_type == 2:
					f=str(worksheet.cell_value(curr_row, curr_cell)).split('.')[1]
					if f == '0':
						cell_value = int(worksheet.cell_value(curr_row, curr_cell))
					else:
						cell_value = worksheet.cell_value(curr_row, curr_cell)
				else:
					cell_value = worksheet.cell_value(curr_row, curr_cell)
				row.append(str(cell_value))
				if curr_cell == num_cells - 1:
					csvfile = open('{}.csv'.format(name), 'a')
					writecsv = csv.writer(csvfile, quoting=csv.QUOTE_ALL)
					writecsv.writerow(row)
					csvfile.close()

def csvtoxls():
	workbook = Workbook(name + '.xls')
	worksheet = workbook.add_worksheet()
	with open(filename, 'r') as f:
		reader = csv.reader(f)
		for r, row in enumerate(reader):
			for c, col in enumerate(row):
				worksheet.write(r, c, col)
	workbook.close()

def csvtoxlsx():
	workbook = Workbook(name + '.xlsx')
	worksheet = workbook.add_worksheet()
	with open(filename, 'r') as f:
		reader = csv.reader(f)
		for r, row in enumerate(reader):
			for c, col in enumerate(row):
				worksheet.write(r, c, col)
	workbook.close()

if args.csv:
	xlstocsv()
elif args.xls:
	csvtoxls()
elif args.xlsx:
	csvtoxlsx()
