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
    spreadsheetDictionary = {}
    for i in range(sheet.nrow):
        currentCell = sheet.cell_value(i,0)
        # At defaults header
        if (currentCell == 'Defaults') and (sheet.cell_value(i, 1) == xlrd.empty_cell.value):
        	spreadsheetDictionary['bladeWidth'] = sheet.cell_value(i+1, 1)
        	spreadsheetDictionary['tolerance'] = sheet.cell_value(i+2, 1)
        	i+=2
        elif (currentCell == 'Panel'):

        elif (currentCell == xlrd.empty_cell.value):
        	continue
