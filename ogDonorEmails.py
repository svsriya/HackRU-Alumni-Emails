#this is the first version created

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import xlrd
import smtplib

#open the excel file
excel_loc = "/Users/svsriya/Documents/PythonPrograms/EmailProgram/TestExcel.xlsx"
workbook = xlrd.open_workbook(excel_loc)
sheet = workbook.sheet_by_index(0)

#variables taken from the excel sheet
num_rows = sheet.nrows
num_cols = sheet.ncols
data = [ [sheet.cell_value(r, c) for c in range( num_cols )] for r in range(num_rows)]

#email login info
your_email = 'svsriya@gmail.com'
temp_pwd = 'dkgrhurtmzsqfgqq'

#cycles through each row in the spreadsheet to get necessary info for email and sends email to each donor
for row in range( 1, num_rows ):
	name = data[row][0]
	amt = data[row][1]
	pos = data[row][2]
	thankyou_note = ("Dear " + name + ", \n" + "\tThank you for your donation of $" + str( "{:1.2f}".format(amt) ) + 
					". With your position as " + pos + ", we are eternally grateful.\n\n")
	
	#create message object instance
	msg = MIMEMultipart()

	#parameters of message
	msg['From'] = your_email
	msg['To'] = your_email
	msg['Subject'] = "Email Application"

	#add in the message body
	msg.attach(MIMEText(thankyou_note, 'plain'))

	#server
	mail = smtplib.SMTP('smtp.gmail.com: 587')
	mail.ehlo()
	mail.starttls()

	#login credentials
	mail.login( your_email, temp_pwd )

	#send message via server
	mail.sendmail( msg['From'], msg['To'], msg.as_string() )

	mail.close()

	print "email sent to %s" %(msg['To'])

