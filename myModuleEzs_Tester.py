#! python3

## Google Sheets
#   SpreadSheet object CRUD
#       Check  SpreadSheet (SS) if exist
#       Create SpreadSheet (SS)
#       Delete SpreadSheet (SS) if exist
#       Copy   SpreadSheet (SS)
#   Sheet object CRUD
#       Check  SpreadSheet SHEET  (SH) if exist
#       Create SpreadSheet SHEET  (SH)
#       Delete SpreadSheet SHEET  (SH) if exist
#       Delete SpreadSheet SHEET  (SH)
#       Copy   SpreadSheet SHEET  (SH)

# define functions to Check, Create and  Delete SS or SH

import ezsheets, sys
import myModuleEzs 

# # SpreadSheet object CRUD
while True:
    break
    ssName = 'Multiple External Sheets'
    #       Check  SpreadSheet (SS) if exist
    if myModuleEzs.checkSpreadSheet(ssName) == True:
        print(f"Yes it exist: {ssName}")
    else:
        print(f"Does not exist: {ssName}")
    ss1 = myModuleEzs.createSpreadSheet(ssName)

    #       Create SpreadSheet (SS)
    ssName = 'Multiple External Sheets2'
    ss2 = myModuleEzs.createSpreadSheet(ssName)
    print(f" New SS..: SS.Title  {str(ss2.title)}  " )

    #       Copy   SpreadSheet (SS1 to SS2)
    ss1[1].updateRow(1, ['Some1', 'data1', 'in', 'the', 'first', 'row'])
    ss1.title='Multiple External Sheets3'
    print("Copied SS" + str(ss1.sheetTitles)) # ss1 before copy 


    #       Delete SpreadSheet (SS) if exist
    if myModuleEzs.deleteSpreadSheet(ssName) == True:
        print(f"Yes I deleted: {ssName}")
    else:
        print(f"Does not exist, I did not delete: {ssName}")
    
    break

#@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    #   Sheet object CRUD
#@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
while True:


    #       Check  SpreadSheet SHEET  (SH) if exist
    ssName = 'Multiple External Sheets'
    shName = 'Spam'
    if myModuleEzs.checkSpreadSheet_Sheet(ssName,shName) == True:
        print(f"Yes Sheet : >>>>{shName}<<<<< Exist in SS {ssName}")
    else:
        print(f"No Sheet : >>>>>{shName}<<<<< Does Not Exist in SS {ssName}")

    #       Create SpreadSheet SHEET  (SH)
    shName = 'NewSh'
    sh1 = myModuleEzs.createSpreadSheet_Sheet(ssName,shName,2)
    print(f"Sheets : >>>>{str(sh1.sheetTitles)}<<<<< Exist in SS ^^^^{ssName}^^^^")

    #       Delete SpreadSheet SHEET  (SH)
    if myModuleEzs.deleteSpreadSheet_Sheet(ssName,shName) == True:
        print(f"Yes I deleted Sheet: >>>>>{shName}<<<<<")
    else:
        print(f"SHEET Does not exist, I did not delete: {shName}")

    #       Copy   SpreadSheet SHEET  (SH)
    ssName = 'Multiple External Sheets'
    ssName2 = 'Multiple External Sheets2'
    ss1 = myModuleEzs.createSpreadSheet(ssName)
    ss2 = myModuleEzs.createSpreadSheet(ssName2)
    ss1[2].updateRow(2, ['Some1', 'data1', 'in', 'the', 'Second', 'row'])
    print(f"B4 Picture  >>>>>{ss2.title}>>>>>  {str(ss2.sheetTitles)} ") # ss2 before copy 
    ss1[2].copyTo(ss2) # Copy the ss1's Spam to the ss2 spreadsheet.
    print(f"After Copy Picture  >>>>>{ss2.title}>>>>>  {str(ss2.sheetTitles)} ") # ss2 AFTER copy 

    sys.exit
    break


