import smtplib
import openpyxl
import sys
# Define todays date
from datetime import datetime, timedelta
now = datetime.now()
# Define what date is in one day
delta = timedelta(days=1)
tomorrow = now + delta
# Convert tomorrow date into string
tomorrow = tomorrow.strftime('%d.%m.%Y')
# Open the excel file with the deadlines on the first sheet
wb = openpyxl.load_workbook('Tasks deadlines.xlsx')
ws = wb.get_sheet_by_name('Sheet1')
# Define the range to be analyzed - in our case column A
colA = ws['A']
# Read the data in column A
for column in ws.columns:
	for colA in column:
# Compare the data in column A to the tomorrow's date
		if colA.value == tomorrow:
# Enter smtp of gmail, login my email, send email, exit smtp
			smtpObj = smtplib.SMTP('smtp.gmail.com', 587)
			smtpObj.starttls()
			smtpObj.login('natalya.agilent@gmail.com', 'YfnfkmzYfnfkmz321')
			smtpObj.sendmail("natalya.agilent@gmail.com", "nakurysheva@mail.ru", "Subject line: Notification. \nHi here is a test message \nThanks \nNatalya")
			smtpObj.quit()
			print("Email sent")
		else:
			print ("Deadline is not today Do not send anything")



