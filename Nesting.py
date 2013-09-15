import xlrd
import Tkinter, tkFileDialog

# This class is the base class for Coupons
class Coupon:
    def __init__(self):
        self.couponSeries
        self.bladeWidth
        self.quantity
    x0 = 0
    y0 = 0
    x1 = 0
    y1 = 0

class Panel:
    def __init__(self, length, width, trimEdge):
        self.length = length
        self.width = width
        self.trimEdge = trimEdge

# This is the extension of Coupon
class RectangularCoupon(Coupon):
    def __init__(self, couponSeries, quantity, length, width, bladeWidth):
        self.couponSeries = couponSeries
        self.quantity = quantity
        self.width = width
        self.length = length
        self.bladeWidth = bladeWidth
        self.area = length*width + (length+width)*bladeWidth #simple area calculation to account for bladeWidth

# Extension of Coupon
class CircularCoupon(Coupon):
    def __init__(self, couponSeries, quantity, radius, bladeWidth):
        self.couponSeries = couponSeries
        self.quantity = quantity
        self.radius = radius
        self.diameter = radius << 1
        self.bladeWidth = bladeWidth
        self.area = self.diameter * self.diameter + (self.diameter + self.diameter)*bladeWidth

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

# Input: anything
# Output: boolean indicating whether s is a number
def isNumber(s):
    try:
        float(s)
        return True
    except ValueError:
        return False

# Input: spreadsheet, row, spreadsheetDictionary
# Output: updated spreadsheetDictionary
def getDefaults(sheet, curr, row, ssd):
    if curr == 'blade width' and isNumber(sheet.cell_value(row,1)):
        ssd['bladeWidth'] = sheet.cell_value(row, 1)
    elif curr == 'tolerance' and isNumber(sheet.cell_value(row,1)):
        ssd['tolerance'] = sheet.cell_value(row, 1)
    elif not isNumber(sheet.cell_value(row,1)):
        return False
    return ssd

# Input: spreadsheet, row, spreadsheetDictionary
# Output: updated spreadsheetDictionary / false
def getPanelDimensions(sheet, row, ssd):
    panelValid = True
    # Check that length, width, and trim edge are all numbers
    for i in range(3):
        if not isNumber(sheet.cell_value(row, i)):
            panelValid = False
    if panelValid:
        pLength = sheet.cell_value(row, 0)
        pWidth = sheet.cell_value(row, 1)
        pTrimEdge = sheet.cell_value(row, 2)
        panel = Panel(pLength, pWidth, pTrimEdge)
        ssd['panel'] = panel
        return ssd
    else:
        return False

# Input: currentCell
# Output: sectionFlag
def getSectionFlag(currentCell):
    if (currentCell == 'Defaults'):
        sectionFlag = 'defaults'
    elif (currentCell == 'Panel'):
        sectionFlag = 'panel'
    elif (currentCell == 'Rectangular Coupons'):
        sectionFlag = 'rectangularCoupons'
    elif (currentCell == 'Circular Coupons'):
        sectionFlag = 'circularCoupons'
    elif (currentCell == xlrd.empty_cell.value):
        sectionFlag = 'break'
    return sectionFlag


# Input: spreadsheet, sectionFlag
# Output: the information for the row parsed
# def extractRowData(sheet, sectionFlag):
#     if sectionFlag

# Input: spreadsheet
# Output: a dictionary containing the spreadsheet information
# This function will create the dictionary containing the blade
# width, tolerance, and information about panels and coupons.
def parseExcel(sheet):
    spreadsheetDictionary = {}
    sectionFlag = 'defaults'
    for i in range(sheet.nrows):
        print "{0} / {1} ".format(i, sheet.nrows)
        currentCell = sheet.cell_value(i,0)
        sectionFlag = getSectionFlag(currentCell)

        print currentCell
        print "spreadsheetDictionary so far: {0}".format(spreadsheetDictionary)
        print "Section flag : {0}".format(sectionFlag)
        raw_input('Check the current cell above')

        continue # Go to next row to avoid having to handle extra bad cases

        if (sectionFlag == 'defaults'):
            defaults = getDefaults(sheet, curr, i, spreadsheetDictionary):
            if defaults: # if valid
                ssd = defaults
            else:
                raise Exception("Defaults must be numbers.")
        elif (sectionFlag == 'break'):
            continue
        elif (sectionFlag == 'panel'):
            panelDim = getPanelDimensions(sheet, i, spreadsheetDictionary)
            if panelDim: # if valid
                ssd = panelDim
            else:
                raise Exception("Panel dimensions must be numbers.")
        elif (sectionFlag == 'rectangularCoupons'): #we are parsing through the data for rectangular/circular coupons
            print "rectangular coupons."
        elif (sectionFlag == 'circularCoupons'):
            print "circular coupons."
    return spreadsheetDictionary









