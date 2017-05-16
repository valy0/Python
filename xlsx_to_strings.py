#!/usr/bin/python
# -*- coding: utf-8 -*-

from openpyxl import load_workbook
import os
import re
import string

"""
This script exports data form .xlsx file to a .strings file.
"""

def main():
	# get .xlsx file data
	xlsx_file_path = raw_input('Path to .strings file (drag and drop will do fine): ')

	# when drag and dropping, there is an extra whitespace at the end => remove it
	if xlsx_file_path[-1:] == ' ':
		xlsx_file_path = xlsx_file_path[:-1]

	# Open the xlsx file and get the first spreadsheet
	wb = load_workbook(filename=xlsx_file_path, read_only=True)
	ws = wb.worksheets[0]

	# Create .strings file
	with open(os.getcwd() + '/' + raw_input('.strings file name: ').strip() + '.strings', 'w+') as strings_file:

		# Loop through the rows of the xlsx file
		for cells in ws.rows:

			# Read and save first 2 cells (tag - translation) to the .strings file
			data = '"' + clean_string(cells[0].value) + '" = "' + clean_string(cells[1].value) + '";\r\n'
			strings_file.write(data)


# 'Cleans' the passed string from characters that should be escaped
def clean_string(dirty_string):
	if dirty_string is None:
		return ' '

	return string.replace(dirty_string.encode('utf-8'), '"', '\\"')


if __name__ == '__main__':
    main()
