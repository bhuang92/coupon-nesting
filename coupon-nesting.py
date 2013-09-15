import xlrd
import Tkinter, tkFileDialog

class Coupon:
    def __init__(self):
        self.couponSeries
        self.bladeWidth
        self.quantity

class RectangularCoupon(Coupon):
    def __init__(self, couponSeries, quantity, length, width, bladeWidth):
        self.couponSeries = couponSeries
        self.quantity = quantity
        self.width = width
        self.length = length
        self.bladeWidth = bladeWidth
        self.area = length*width + (length+width)*bladeWidth #simple area calculation to account for bladeWidth
        self.x0
        self.y0
        self.x1
        self.y1

class CircularCoupon(Coupon):
    def __init__(self, couponSeries, quantity, radius, bladeWidth):
        self.couponSeries = couponSeries
        self.quantity = quantity
        self.radius = radius
        self.diameter = radius << 1
        self.bladeWidth = bladeWidth
        self.area = self.diameter * self.diameter + (self.diameter + self.diameter)*bladeWidth
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
# def extractRowData(sheet, sectionFlag):
#     if sectionFlag
# Input: spreadsheet
# Output: a dictionary containing the spreadsheet information
# This function will create the dictionary containing the trim edge, blade
# width, tolerance, 
def parseExcel(sheet):
    spreadsheetDictionary = {}
    sectionFlag = 'defaults'
    for i in range(sheet.nrows):
        print "{0} / {1} ".format(i, sheet.nrows)
        currentCell = sheet.cell_value(i,0)
        print currentCell
        raw_input('Check the current cell above')
        # At defaults header
        if (sectionFlag == 'defaults'):
        	spreadsheetDictionary['bladeWidth'] = sheet.cell_value(i+1, 1)
        	spreadsheetDictionary['tolerance'] = sheet.cell_value(i+2, 1)
        	i+=2
        elif (currentCell == 'Panel'):
            sectionFlag = 'panel'
            pLength = sheet.cell_value(i+1, 0)
            pWidth = sheet.cell_value(i+1, 1)
            pTrimEdge = sheet.cell_value(i+1, 2)
            panelDimensions = {'pLength':pLength, 'pWidth':pWidth, 'pTrimEdge':pTrimEdge}
            spreadsheetDictionary['panelDimensions'] = panelDimensions
            i+=1
        elif (currentCell == 'Rectangular Coupons'):
            sectionFlag = 'rectCoupons'
        elif (currentCell == "Circiular Coupons"):
            sectionFlag = 'circularCoupons'
        elif (currentCell == xlrd.empty_cell.value):
        	continue
        else: #we are parsing through the data for rectangular/circular coupons
            #extractRowData(sheet, sectionFlag)
            print "Currently done."

        return spreadsheetDictionary



