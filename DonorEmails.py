from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import xlrd
import smtplib

#email login info
your_email = 'randomgmail'
temp_pwd = 'password'

def create_donor_list( file_loc ):
	workbook = xlrd.open_workbook(file_loc)
	sheet = workbook.sheet_by_index(0)

	#variables taken from the excel sheet
	num_rows = sheet.nrows
	num_cols = sheet.ncols
	dlist = [ [sheet.cell_value(r, c) for c in range( num_cols )] for r in range(num_rows)]
	return dlist

def custom_ty_note( name, amt, pos ):
	thankyou_note = ("Dear " + name + ", \n" + "\tThank you for your donation of $" + str( "{:1.2f}".format(amt) ) + 
						". With your position as " + pos + ", we are eternally grateful.\n\n")
	return thankyou_note


def send_email( to_email, from_email, body ):
	#create message object instance

	msg = MIMEMultipart()

	#parameters of message
	msg['From'] = from_email
	msg['To'] = to_email
	msg['Subject'] = "Email Application"

	#add in the message body
	msg.attach(MIMEText(body, 'plain'))

	#server
	mail = smtplib.SMTP('smtp.gmail.com: 587')
	mail.ehlo()
	mail.starttls()

	#login credentials
	mail.login( your_email, temp_pwd )

	#send message via server
	mail.sendmail( msg['From'], msg['To'], msg.as_string() )

	mail.close()

	return "email sent to %s" %(msg['To'])

def main():
	#open the excel file
	excel_loc = "location of excel file"

	#create list of donors
	data = create_donor_list( excel_loc )

	#cycles through each row in the spreadsheet to get necessary info for email and sends email to each donor
	for row in range( 1, len(data) ):
		name = data[row][0]
		amt = data[row][1]
		pos = data[row][2]
		email = 'randomgmail'

		email_body = custom_ty_note( name, amt, pos )
		print send_email( email, your_email, email_body )

if __name__ == '__main__':
   main()		

