# Create Google-Sheet from Excel 
#        get ENTIRE Row
#        Update  ENTIRE Row
#        get ENTIRE Column 
#        Update  ENTIRE Column
import ezsheets, sys
ss = ezsheets.upload('C:/Users/ben/OneDrive/Documents/Gaba_Docs/development/Python/Python_Ben/pgmInput/produceSales.xlsx')
sheet = ss[0]
sheet.getRow(1) # The first row is row 1, not row 0.
#['PRODUCE', 'COST PER POUND', 'POUNDS SOLD', 'TOTAL', '', '']
sheet.getRow(2)
#['Potatoes', '0.86', '21.6', '18.58', '', '']
columnOne = sheet.getColumn(1)
sheet.getColumn(1)
#['PRODUCE', 'Potatoes', 'Okra', 'Fava beans', 'Watermelon', 'Garlic',

sheet.getColumn('A') # Same result as getColumn(1)
# ['PRODUCE', 'Potatoes', 'Okra', 'Fava beans', 'Watermelon', 'Garlic',

sheet.getRow(3)
# ['Okra', '2.26', '38.6', '87.24', '', '']
sheet.updateRow(3, ['Pumpkin', '11.50', '20', '230'])
sheet.getRow(3)
# ['Pumpkin', '11.50', '20', '230', '', '']
columnOne = sheet.getColumn(1)
# for i, value in enumerate(columnOne):
for i, value in enumerate(columnOne):
    # Make the Python list contain uppercase strings:
    columnOne[i] = value.upper()
    sheet.updateColumn(1, columnOne) # Update the entire column in one
    if i > 10:
        break
sys.exit