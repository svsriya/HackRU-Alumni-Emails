from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import xlrd
import smtplib

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
		return 'email: ' + self.email + ', causes: ' + str( self.causes )


def create_donor_list( file_loc ):

	workbook = xlrd.open_workbook( file_loc )
	sheet = workbook.sheet_by_index( 0 )

	# variables taken from the excel sheet
	num_rows = sheet.nrows
	num_cols = sheet.ncols
	plp_dict = dict()

	for row in range ( 1, num_rows ):
		nm = sheet.cell_value( row, 0 )
		cs = sheet.cell_value( row, 1 )
		em = sheet.cell_value( row, 2 )

		if nm in plp_dict.keys():
			if cs in plp_dict[nm].causes:
				continue
			else:
				plp_dict[nm].add_cause( cs )

		elif em.strip():
			css = []
			new_entry = Person( em, css )
			new_entry.add_cause( cs )
			plp_dict[nm] = new_entry

	return plp_dict

def custom_ty_note( name, cauz ):

	thankyou_note = "Dear " + name + ", \n\n" + "\tThank you for your donation. You have donated to the following causes: \n" 

	for c in cauz:
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

	# email login info
	your_email = 'svsriya@gmail.com'
	temp_pwd = 'derpdederpdederp'

	# open the excel file
	excel_loc = "/Users/svsriya/Documents/PythonPrograms/EmailProgram/Practice.xlsx"

	# create list of donors
	data = create_donor_list( excel_loc ) 	

	# cycles through each row in the spreadsheet to get necessary info for email and sends email to each donor
	
	for key in data:
		name = key
		email = data[key].email
		cauz = data[key].causes

		email_body = custom_ty_note( name, cauz )
		print email_body + "\n"
		print send_email( email, your_email, email_body, temp_pwd )


if __name__ == '__main__':
   main()		

