# -*- coding: utf-8 -*-

import conf
import openpyxl

# configuration
INPUT_FILE_PATH = conf.INPUT_FILE_PATH
START_CELL = 2

# open input xls table
wb = openpyxl.load_workbook(INPUT_FILE_PATH)

# load sheets
sheets = wb.get_sheet_names()
if len(sheets)>0:
	sheets = wb.get_sheet_by_name(
		sheets[0])
else:
	raise Exception(
		"No sheets can be found!")


# iteration and update

detail = ''
description = ''
initial_value = sheets['B'+str(
	START_CELL)].value + ' ' + str(sheets['D'+str(
		START_CELL)].value) + u'\u20ac'
is_start = True
acc = 0

for idx in range(1, sheets.max_row):
	inv_value = sheets['A'+str(idx + 1)].value
	if inv_value is None:
		acc += 1
		description = ', ' +sheets['B'+str(
			idx+1)].value + ' ' + str(sheets['D'+str(
			idx+1)].value) + u'\u20ac'
		sheets['J'+str(START_CELL)] = sheets['J'+str(START_CELL)].value + description
	else:
		if is_start == True:
		    sheets['J'+str(START_CELL)] = initial_value + description
		    is_start = False
		else:
			START_CELL = START_CELL + 1 + acc
			initial_value = sheets['B'+str(
				idx+1)].value + ' ' + str(sheets['D'+str(
					idx+1)].value) + u'\u20ac'
			sheets['J'+str(START_CELL)] = initial_value

		acc = 0

		description = ''
		

wb.save(INPUT_FILE_PATH)















