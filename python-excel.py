import xlrd
import csv

def ask_for_filepath():
    return input('Excel file path: ')

def ask_for_sheetname():
    return input('Sheet name: ')

def ask_for_output_filename():
    return input('Output file name: ')

def ask_for_column_index():
    return input('Column Index: ')

def ask_for_pattern():
    return input('Pattern to be found: ')

def create_list_based_on_pattern(pattern, column_index):
    return [sheet.cell_value(i, column_index) for i in range(sheet.nrows) if str(sheet.cell_value(i, column_index)).startswith(str(pattern))]


filepath = ask_for_filepath()
workbook = xlrd.open_workbook(filepath)
sheetname = ask_for_sheetname()
sheet = workbook.sheet_by_name(sheetname)

column_index = ask_for_column_index()
pattern = ask_for_pattern()
pattern_list = create_list_based_on_pattern(pattern, column_index)

print(pattern_list)

output_file = ask_for_output_filename()
with open(output_file, 'w') as csv_file:
    writer = csv.writer(csv_file, delimiter=',', quotechar='"')
    writer.writerows([pattern_list])

