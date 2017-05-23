import smtplib
import openpyxl
import sys
# Open the excel file with the deadlines on the first sheet
wb = openpyxl.load_workbook('Tasks deadlines.xlsx')
ws = wb.get_sheet_by_name('Sheet1')
c = ws.cell(row=2, column=2)
from datetime import datetime, timedelta
now = datetime.now()
delta = timedelta(days=1)
tomorrow = now + delta
tomorrow = tomorrow.strftime('%d.%m.%Y')
if c.value == tomorrow:
	print ("Deadline is tomorrow Send the email")
else:
	print ("Deadline is not today Do not send anything")

	#Не работает условие, так как не работает отображение значения ячейки