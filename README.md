# GmailToDrive

#### This is a Google Apps Scripts comprised of three functions: GmailScan(), convertExceltoGS() and CopyRange().

#### GmailScan() loops through the Gmail App searching for emails from the specific sender and saves the latest message's excel attachment then calls the function convertExceltoGS() to change file format from .xlsx to google sheet.

#### convertExceltoGS() does as mentioned previously - converts excel file to google sheet and appends the email message's date to the new file name. Lastly, it calls function CopyRange() to paste contents of newly originated google sheet to the master file in Google Drive.

#### CopyRange() pulls the contents of the newly created google sheet and appends the new data to existing rows in master file. It ensures the data is pasted in next available blank row, dropping contents in correct columns.
