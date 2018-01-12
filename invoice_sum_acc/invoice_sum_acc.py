import conf
import openpyxl

# configuration
INPUT_FILE_PATH = conf.INPUT_FILE_PATH

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

amount_sum = 0
start_cell = 0
end_cell = 0
start_number = 1

# iteration and update

for idx in range(1,sheets.max_row):
	inv_value = sheets['A'+str(idx + 1)].value
	if inv_value is None:
		# set sum of money
		end_cell = idx + 1
		sheets['G'+str(start_cell)] = "=SUM({}:{})".format(
			'D'+str(start_cell), 'D'+str(
				end_cell))
		# set start column
		start_number += 1
		sheets['H'+str(idx+1)] = start_number
	else:
		start_number = 1
		start_cell = idx + 1
		sheets['G'+str(idx+1)] = sheets['D'+str(
			idx+1)].value
		sheets['H'+str(idx+1)] = start_number

wb.save(INPUT_FILE_PATH)