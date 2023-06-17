#! python3
#
# Creating, Uploading, and Listing Spreadsheets 
import ezsheets, sys
import myModulePath

outputPath = str(myModulePath.allPath['logPath']['All']['comment'])
# From spreadsheetId
ss = ezsheets.Spreadsheet('1QLN4CqSiZB-eDHwCcDJhrSiU8urZQDtkoVKwHt-YRZg')
print("Create from spreadsheetId")
print(str(ss))
# Spreadsheet(spreadsheetId='1QLN4CqSiZB-eDHwCcDJhrSiU8urZQDtkoVKwHt-YRZg')
print(str(ss.title))
#  'Education Data'

# From Full URL
ss = ezsheets.Spreadsheet('https://docs.google.com/spreadsheets/d/1QLN4CqSiZB-eDHwCcDJhrSiU8urZQDtkoVKwHt-YRZg/edit#gid=0')
print("Create from Full URL")
print(str(ss))
# Spreadsheet(spreadsheetId='1QLN4CqSiZB-eDHwCcDJhrSiU8urZQDtkoVKwHt-YRZg')
print(str(ss.title))
#  'Education Data'

# From Full name
ss = ezsheets.Spreadsheet('ezs_Education Data')
print("Create from filename")
print(str(ss))
# Spreadsheet(spreadsheetId='1QLN4CqSiZB-eDHwCcDJhrSiU8urZQDtkoVKwHt-YRZg')
print(str(ss.title))
#  'Education Data'

#To upload an existing Excel, OpenOffice, CSV, or TSV spreadsheet to Google Sheets
ss = ezsheets.upload('duesRecords.xlsx')
ss.title='ezs_DueRecords'
print("Upload from Excel")
print(str(ss))
# Spreadsheet(spreadsheetId='1QLN4CqSiZB-eDHwCcDJhrSiU8urZQDtkoVKwHt-YRZg')
print(str(ss.title))
#  'Education Data'

#list SpreadSheets
ssList = ezsheets.listSpreadsheets()
print("LIST All google spreadsheets")
for i, filename in enumerate(ssList):
    # print name of items
    print(str(ssList[filename]))
    break

# Download to PDF,XLSX, CSV, HTML, TSV (Tab Seperated Variables)
# downloads to CWD
ss = ezsheets.Spreadsheet('Ch14_Get_Name_email (Responses)')
Pss.downloadAsExcel() # Downloads the spreadsheet as an Excel file.
ss.downloadAsODS() # Downloads the spreadsheet as an OpenOffice file.
ss.downloadAsCSV() # Only downloads the first sheet as a CSV file.
ss.downloadAsTSV() # Only downloads the first sheet as a TSV file.
ss.downloadAsPDF() # Downloads the spreadsheet as a PDF.
ss.downloadAsHTML() # Downloads the spreadsheet as a ZIP of HTML files.

sys.exit

