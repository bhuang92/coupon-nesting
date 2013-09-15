'''
Author: Basil Huang
Date: September 2013
'''

import xlrd
import Tkinter, tkFileDialog
import copy

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
def oss():#openSpreadsheet():
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

# Input: currentCell
# Output: sectionFlag
def getSectionFlag(currentSF, currentCell):
    sf = currentSF
    if (currentCell == 'Defaults'):
        print "1"
        sf = 'defaults'
    elif (currentCell == 'Panel'):
        sf = 'panel'
        print "2"
    elif (currentCell == 'Rectangular Coupons'):
        sf = 'rectangularCoupons'
        print "3"
    elif (currentCell == 'Circular Coupons'):
        sf = 'circularCoupons'
        print "4"
    elif (currentCell == xlrd.empty_cell.value):
        sf = 'break'
        print "5"
    return sf

# Input: spreadsheet, row, spreadsheetDictionary
# Output: updated spreadsheetDictionary
def getDefaults(sheet, curr, row, ssd):
    print "sheet: {0}".format(sheet)
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

# Input: spreadsheet, sectionFlag
# Output: the information for the row parsed
#def extractRowData(sheet, row, ssd):
 #   if sectionFlag

# Input: spreadsheet
# Output: a dictionary containing the spreadsheet information
# This function will create the dictionary containing the blade
# width, tolerance, and information about panels and coupons.
def parseExcel(sheet):
    spreadsheetDictionary = {}
    rectangularCouponArray = []
    circularCouponArray = []
    sectionFlag = 'defaults'
    for i in range(sheet.nrows):
        print "{0} / {1} ".format(i, sheet.nrows)
        currentCell = sheet.cell_value(i,0)
        sectionFlag = getSectionFlag(sectionFlag, currentCell)

        print currentCell
        print "spreadsheetDictionary so far: {0}".format(spreadsheetDictionary)
        print "Section flag : {0}".format(sectionFlag)

        if (sectionFlag == 'defaults'):
            print "Got here 1"
            defaults = getDefaults(sheet, curr, i, spreadsheetDictionary)
            if defaults: # if valid
                spreadsheetDictionary = defaults
            else:
                raise Exception("Defaults must be numbers.")
        elif (sectionFlag == 'break'):
            print "got here 2"
            continue
        elif (sectionFlag == 'panel'):
            print "got here 3"
            panelDim = getPanelDimensions(sheet, i, spreadsheetDictionary)
            if panelDim: # if valid
                spreadsheetDictionary = panelDim
            else:
                raise Exception("Panel dimensions must be numbers.")
        elif (sectionFlag == 'rectangularCoupons'): #we are parsing through the data for rectangular/circular coupons
            print "rectangular coupons."
        elif (sectionFlag == 'circularCoupons'):
            print "circular coupons."

        raw_input('Check the current cell above')
    return spreadsheetDictionary









