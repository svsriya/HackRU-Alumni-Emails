from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import xlrd
import smtplib

# email login info
your_email = 'random email'
temp_pwd = 'random password'

# open the excel file
excel_loc = "location of excel"

class Person:

	# this makes it easier to handle people's names who appear multiple times on the excel
	def __init__( self, email, causes ):
		self.email = email
		self.causes = []

	def add_cause( self, cause ):
		self.causes.append( cause )

	def __repr__( self ):
		return 'email: ' + self.email + ', causes: ' + str( self.causes )

	def __str__( self ):
		return self.__repr__()


def create_donor_list( file_loc ):

	# variables taken from the excel sheet
	workbook = xlrd.open_workbook( file_loc )
	sheet = workbook.sheet_by_index( 0 )

	num_rows = sheet.nrows
	num_cols = sheet.ncols
	people_dict = dict()

	for row in range ( 1, num_rows ):
		name = sheet.cell_value( row, 0 )
		cause = sheet.cell_value( row, 1 )
		email = sheet.cell_value( row, 2 )

		if name in people_dict.keys():
			if cause in people_dict[name].causes:
				continue
			else:
				people_dict[name].add_cause( cause )

		elif email.strip():
			causes_list = []
			new_entry = Person( email, causes_list )
			new_entry.add_cause( cause )
			people_dict[name] = new_entry

	return people_dict

def custom_ty_note( name, causes_list ):

	thankyou_note = "Dear " + name + ", \n\n" + "\tThank you for your donation. You have donated to the following cause(s): \n" 

	for c in causes_list:
		thankyou_note += "\n\t" + c

	return thankyou_note


def send_email( to_email, from_email, password, body ):

	# create message object instance
	msg = MIMEMultipart()

	# parameters of message
	msg['From'] = from_email
	msg['To'] = to_email
	msg['Subject'] = "Email Application"

	# add in the message body
	msg.attach( MIMEText( body, 'plain' ) )

	# server
	mail = smtplib.SMTP( 'smtp.gmail.com: 587' )
	mail.ehlo()
	mail.starttls()

	# login credentials
	mail.login( from_email, password )

	# send message via server
	mail.sendmail( msg['From'], msg['To'], msg.as_string() )

	mail.close()

	return "email sent to %s" % msg['To']

def main():

	# create list of donors
	data = create_donor_list( excel_loc ) 	

	# cycles through each row in the spreadsheet to get necessary info for email and sends email to each donor
	
	for key in data:
		name = key
		email = data[key].email
		causes_list = data[key].causes

		email_body = custom_ty_note( name, causes_list )
		print email_body + "\n"
		print send_email( email, your_email, temp_pwd, email_body )


if __name__ == '__main__':
   main()		

