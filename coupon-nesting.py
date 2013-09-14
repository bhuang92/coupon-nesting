import xlrd
import Tkinter, tkFileDialog

# Input: none
# Output: spreadsheet 
# This function uses the Tkinter library to ask the user for the location
# of the spreadsheet. Then the spreadsheet is opened up and the spreadsheet 
# object is returned.
def openSpreadsheet():
	root = Tkinter.Tk()
	root.withdraw()
	file_path = tkFileDialog.askopenfilename()
	book = xlrd.open_workbook(file_path)
    return book.sheet_by_index(0)

# Input: spreadsheet
# Output: a dictionary containing the spreadsheet information
# This function will create the dictionary containing the trim edge, blade
# width, tolerance, 
def parseExcel(sheet):
    spreadsheetDictionary = {'trimEdge': 1, 'bladeWidth': 0.07, 'tolerance': 0.015}
    for i in range(18):
        currentCell = sh.cell_value(i,0)





