#! python3

# # SpreadSheet object CRUD
#   Check  SpreadSheet (SS) if exist
#   Create SpreadSheet (SS)
#   Delete SpreadSheet (SS) if exist
#   Copy   SpreadSheet (SS)
# Sheet object CRUD
#        Check  SpreadSheet SHEET  (SH) if exist
#        Create SpreadSheet SHEET  (SH)
#        Delete SpreadSheet SHEET  (SH) if exist
#        Delete SpreadSheet SHEET  (SH)
#        Copy   SpreadSheet SHEET  (SH)

# define functions to Check, Create and  Delete SS or SH

import ezsheets, sys


""" @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
                    Spreadsheet Object
@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ """

def checkSpreadSheet(name):
    shList= ezsheets.listSpreadsheets()
    for i, value in enumerate(shList):
        # print spreadsheet name
        if shList[value]==name:
            return True
    return False

def createSpreadSheet(ssName):
    if checkSpreadSheet(ssName) == True:
        ss=ezsheets.Spreadsheet(ssName)             # create Existing
    else:
        ss = ezsheets.createSpreadsheet(ssName)     # create New
    return ss

def deleteSpreadSheet(ssName):
    if checkSpreadSheet(ssName) == True:
        ss=ezsheets.Spreadsheet(ssName)             # delete
        ss.delete(permanent= True)
        return True
    else:
        return False

""" @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
                    Spreadsheet SHEET'S Object
@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ """

def checkSpreadSheet_Sheet(ssName,shName):    
    if checkSpreadSheet(ssName) == True:
        sh= ezsheets.Spreadsheet(ssName)
        shList = sh.sheetTitles #  list names of sheets
        for i, value in enumerate(shList):
            # print spreadsheet name
            if value==shName:
                return True
        return False
    else:
        return False
    
def getSpreadSheetName_Sheet(ssName,shIndex=0):    
    if checkSpreadSheet(ssName) == True:
        sh= ezsheets.Spreadsheet(ssName)
        shList = sh.sheetTitles #  list names of sheets
        for i, value in enumerate(shList):
            # print spreadsheet name
            if i==shIndex:
                return str(value)
        return " "
    else:
        return " "

# In the function createSpreadSheet_Sheet, you’ve added the default value -1 to the parameter shIndex. 
# This doesn’t mean that the value of shIndex will always be -1. 
# If you pass an argument corresponding to shIndex when you call the function, 
# then that argument will be used as the value for the parameter. However, 
# if you don’t pass any argument, then the default value will be used.
def createSpreadSheet_Sheet(ssName,shName,shIndex=-1):
    if checkSpreadSheet_Sheet(ssName,shName) == True:
        ss= ezsheets.Spreadsheet(ssName)
        ss[shName].delete()  # Delete sheet object
        ss.refresh()        # Refsresh Spreadsheet object
    else:
        ss= ezsheets.Spreadsheet(ssName)
    if shIndex > -1:
        ss.createSheet(shName,shIndex ) 
    else:
        ss.createSheet(shName ) 
    return ss

def deleteSpreadSheet_Sheet(ssName,shName,shIndex=-1):
    if checkSpreadSheet_Sheet(ssName,shName) == True:
        ss= ezsheets.Spreadsheet(ssName)
        ss[shName].delete()  # Delete sheet object
        ss.refresh()         # Refsresh Spreadsheet object
        return True
    else:
        return False


sys.exit