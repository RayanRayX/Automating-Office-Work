import openpyxl

def create_excel_file():
	workbook_name = input("Enter Name of Your files: ")

	workbook_name += '.xlsx'

	worksheet_name = input("Enter Name of worksheet: ")

	num_columns = int(input("Enter the Number of column you want: "))

	workbook = openpyxl.Workbook()
	
	first_worksheet = workbook.active
	first_worksheet.title = worksheet_name

	column_header = [input(f"Enter header for column {i}: ") for i in range(1, num_columns + 1)]

	for col_num, header in enumerate(column_header, 1):
		first_worksheet.cell(row=1, column = col_num, value=header)
	
	workbook.save(workbook_name)

	workbook.close()

	print("Done")
	
create_excel_file()