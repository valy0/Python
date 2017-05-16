#!/usr/bin/python
# -*- coding: utf-8 -*-

from openpyxl import Workbook
import os
import re
import string

"""
This script exports data form .strings file to an excel spreadsheet file.
"""

def main():
	data = ''

	# get .strings file data
	strings_file_path = raw_input('Path to .strings file (drag and drop will do fine): ')

	# when drag and dropping, there is an extra whitespace at the end => remove it
	if strings_file_path[-1:] == ' ':
		strings_file_path = strings_file_path[:-1]

	# open and read data from provided file
	with open(strings_file_path) as strings_file:
		data = strings_file.read()

	# remove any unnecessary info from file
	clean_data = re.findall(r'".+"\s+?=\s+?".+?";', data)

	# create xlsx file
	wb = Workbook()

	# set worksheet
	ws = wb.active

	for row in clean_data:
		# split row by '=' to separate key and value
		components = re.split(r'\s+?=\s+?', row)

		# get key tag and remove '"' from start and end
		tag = components[0][1:][:-1]

		# get translation value and remove '"' from start and '";' from end
		translation = components[len(components) - 1][1:][:-2]

		# add data to xlsx file
		ws.append([tag, translation])

	# save xlsx file
	wb.save(os.getcwd() + '/' + raw_input('Spreadsheet file name: ').strip() + '.xlsx')


# 'Cleans' the passed string from characters that should be escaped
def clean_string(dirty_string):
	if dirty_string is None:
		return ' '

	return string.replace(dirty_string, '\\"', '"')


if __name__ == '__main__':
    main()
