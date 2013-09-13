import xlrd

def parseExcel():
	#First read in inputs from the excel file
	book = xlrd.open_workbook("../Spreadsheets/example.xlsx")
	