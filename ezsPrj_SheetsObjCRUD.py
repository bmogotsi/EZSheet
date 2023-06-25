#! python3

# # Sheet object CRUD
#        Create Sheet
#        Update Sheet
#        Delete Sheet 
#        Copy Sheet

# define functions to Check, Create and  Delete SHEETS in a spread sheet 
#?????????????????????????????????????????????????????????????????????????????????????????????????????????
# challenge can you define them as modules so that they are external and available to other scripts 
#?????????????????????????????????????????????????????????????????????????????????????????????????????????

import ezsheets, sys
def checkSpreadSheet(name):
    shList= ezsheets.listSpreadsheets()
    for i, value in enumerate(shList):
        # print spreadsheet name
        if shList[value]==name:
            return True
    return False


def checkSpreadSheet_Sheet(ssStr,shStr):    
    if checkSpreadSheet(ssStr) == True:
        sh= ezsheets.Spreadsheet(ssStr)
        shList = sh.sheetTitles #  list names of sheets
        for i, value in enumerate(shList):
            # print spreadsheet name
            if value==shStr:
                return True
        return False
    else:
        return False

# In the function createSpreadSheet_Sheet, you’ve added the default value -1 to the parameter shIndex. 
# This doesn’t mean that the value of shIndex will always be -1. 
# If you pass an argument corresponding to shIndex when you call the function, 
# then that argument will be used as the value for the parameter. However, 
# if you don’t pass any argument, then the default value will be used.
def createSpreadSheet_Sheet(ssStr,shStr,shIndex=-1):
    if checkSpreadSheet_Sheet(ssStr,shStr) == True:
        # ss.delete(shStr,permanent=True) this will delete the actual sheet name
        ss[shStr].delete()  # Delete sheet object
        ss.refresh()        # Refsresh Spreadsheet object
    if shIndex > -1:
        ss.createSheet(shStr,shIndex ) 
    else:
        ss.createSheet(shStr ) 

#        Create Sheet named "Multiple Sheets"
ssName = 'Multiple Sheets'

if checkSpreadSheet(ssName) == True:
    ss=ezsheets.Spreadsheet(ssName)
else:
    ss = ezsheets.createSpreadsheet(ssName)  # "Sheet1" is defaulted

shName = 'Spam'
createSpreadSheet_Sheet(ssName,shName)
shName = 'Eggs'
createSpreadSheet_Sheet(ssName,shName) # Create another new sheet.
shName = 'Bacon'
createSpreadSheet_Sheet(ssName,shName,2) # Create a sheet at index 0 (means b4 index 0) in the list of sheets.
# <Sheet sheetId=814694991, title='Bacon', rowCount=1000, columnCount=26>

pTitles= ss.sheetTitles
print(f"Sheet names:  {pTitles} and IndexOfSheet Eggs and  {str(ss.sheetTitles.index('Eggs'))} ")
print(*pTitles)


sys.exit