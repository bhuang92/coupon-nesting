'''
Author: Basil Huang
Date: September 2013
'''

import xlrd
import Tkinter as tk, tkFileDialog
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
        self.diameter = radius *2
        self.bladeWidth = bladeWidth
        self.area = self.diameter * self.diameter + (self.diameter + self.diameter)*bladeWidth

# Input: none
# Output: spreadsheet 
# This function uses the Tkinter library to ask the user for the location
# of the spreadsheet. Then the spreadsheet is opened up and the spreadsheet 
# object is returned.
def oss():#openSpreadsheet():
    root = tk.Tk()
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
        sf = 'defaults'
    elif (currentCell == 'Panel'):
        sf = 'panel'
    elif (currentCell == 'Rectangular Coupons'):
        sf = 'rectangularCoupons'
    elif (currentCell == 'Circular Coupons'):
        sf = 'circularCoupons'
    elif (currentCell == xlrd.empty_cell.value):
        sf = 'break'
    return sf

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

# Input: spreadsheet, row, sectionFlag
# Output: a rectangularCoupon
def getRectangularCoupon(sheet, row, ssd):
    series = sheet.cell_value(row, 0)
    quantity = sheet.cell_value(row, 1)
    length = sheet.cell_value(row, 2)
    width = sheet.cell_value(row, 3)
    bladeWidth = sheet.cell_value(row, 4)
    if bladeWidth == '':
        bladeWidth = ssd['bladeWidth']
    if isNumber(quantity) and isNumber(length) and isNumber(width):
        rectCoupon = RectangularCoupon(series, quantity, length, width, bladeWidth)
        return rectCoupon
    else:
       return False

# Input: spreadsheet, row, sectionFlag
# Output: a circularCoupon
def getCircularCoupon(sheet, row, ssd):
    series = sheet.cell_value(row, 0)
    quantity = sheet.cell_value(row, 1)
    radius = sheet.cell_value(row, 2)
    bladeWidth = sheet.cell_value(row, 3)
    if bladeWidth == '':
        bladeWidth = ssd['bladeWidth']
    if isNumber(quantity) and isNumber(radius):
        circCoupon = CircularCoupon(series, quantity, radius, bladeWidth)
        return circCoupon
    else:
       return False

# Input: spreadsheet
# Output: a dictionary containing the spreadsheet information
# This function will create the dictionary containing the blade
# width, tolerance, and information about panels and coupons.
def parseExcel(sheet):
    spreadsheetDictionary = {}
    rectangularCouponArray = []
    circularCouponArray = []
    sectionFlag = ''
    for i in range(sheet.nrows):
        #print "{0} / {1} ".format(i, sheet.nrows)
        currentCell = sheet.cell_value(i,0)

        # If we move into a new section, then the sectionFlag will change.
        # This will not move the iterator and will go into the parsing phase below.
        # Use a continue if the flags are not equal and update sectionFlag.
        newSectionFlag = getSectionFlag(sectionFlag, currentCell)
        if not (newSectionFlag == sectionFlag):
            sectionFlag = newSectionFlag
            continue

        # print currentCell
        # print "spreadsheetDictionary so far: {0}".format(spreadsheetDictionary)
        # print "Section flag : {0}".format(sectionFlag)

        if (sectionFlag == 'defaults'):
            defaults = getDefaults(sheet, currentCell, i, spreadsheetDictionary)
            if defaults: # if valid
                spreadsheetDictionary = defaults
            else:
                continue
        elif (sectionFlag == 'break'):
            continue
        elif (sectionFlag == 'panel'):
            panelDim = getPanelDimensions(sheet, i, spreadsheetDictionary)
            if panelDim: # if valid
                spreadsheetDictionary = panelDim
            else:
                continue
        elif (sectionFlag == 'rectangularCoupons'): #we are parsing through the data for rectangular/circular coupons
            if currentCell == 'Coupon series':
                continue
            else:
                rectCoupon = getRectangularCoupon(sheet, i, spreadsheetDictionary)
                if rectCoupon: # if valid
                    rectangularCouponArray.append(rectCoupon)
                else:
                    continue
        elif (sectionFlag == 'circularCoupons'):
            if currentCell == 'Coupon series':
                continue
            else:
                circCoupon = getCircularCoupon(sheet, i, spreadsheetDictionary)
                if circCoupon: # if valid
                    circularCouponArray.append(circCoupon)
                else:
                    continue
        spreadsheetDictionary['rectCoupons'] = rectangularCouponArray
        spreadsheetDictionary['circCoupons'] = circularCouponArray
    return spreadsheetDictionary

''' Excel parsing code is above '''
''' Algorithm code is below '''

''' Drawing code is below '''
import Image, ImageDraw

width = 400
height = 300
center = height//2
white = (255, 255, 255)
green = (0,128,0)

root = tk.Tk()

# Tkinter create a canvas to draw on
cv = tk.Canvas(root, width=width, height=height, bg='white')
cv.pack()

# PIL create an empty image and draw object to draw on
# memory only, not visible
image1 = Image.new("RGB", (width, height), white)
draw = ImageDraw.Draw(image1)

# do the Tkinter canvas drawings (visible)
cv.create_line([0, center, width, center], fill='green')

# do the PIL image/draw (in memory) drawings
draw.line([0, center, width, center], green)

# PIL image can be saved as .png .jpg .gif or .bmp file (among others)
filename = "my_drawing.png"
image1.save(filename)

root.mainloop()






