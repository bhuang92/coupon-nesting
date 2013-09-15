import xlrd
import Tkinter, tkFileDialog

class RectangularCoupon:
    def __init__(self):
        self.couponSeries
        self.width
        self.length
        self.quantity
        self.bladeWidth
        self.area
        self.x0
        self.y0
        self.x1
        self.y1


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

# Input: spreadsheet, sectionFlag
# Output: the information for the row parsed
def extractRowData(sheet, sectionFlag):

# Input: spreadsheet
# Output: a dictionary containing the spreadsheet information
# This function will create the dictionary containing the trim edge, blade
# width, tolerance, 
def parseExcel(sheet):
    spreadsheetDictionary = {}
    sectionFlag = 'defaults'
    for i in range(sheet.nrow):
        currentCell = sheet.cell_value(i,0)
        # At defaults header
        if (sectionFlag == 'defaults'):
        	spreadsheetDictionary['bladeWidth'] = sheet.cell_value(i+1, 1)
        	spreadsheetDictionary['tolerance'] = sheet.cell_value(i+2, 1)
        	i+=2
        elif (currentCell == 'Panel'):
            sectionFlag = 'panel'
            continue
        elif (currentCell == 'Rectangular Coupons'):
            sectionFlag = 'rectCoupons'
        elif (currentCell == "Circiular Coupons"):
            sectionFlag = 'circularCoupons'
        elif (currentCell == xlrd.empty_cell.value):
        	continue
        else:
            extractRowData(sheet, sectionFlag)



