import pandas as pd
import openai
import re
from unidecode import unidecode
import concurrent.futures
import time, os
import csv
import json
import email, imaplib
from enum import Enum
import pprint
from email.header import decode_header
import winreg
import ctypes
import subprocess
import sys
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime, timedelta
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import PyPDF2

from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import WebDriverException
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import ElementClickInterceptedException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import ElementNotInteractableException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait, Select
from bs4 import BeautifulSoup

smtp_server = "smtp.gmail.com"
email_username = "5Cavailabilities@gmail.com"
email_app_pass = "nigelshamash1"

result_array = []

class OrderedEnum(Enum):
	def __ge__(self, other):
		if self.__class__ is other.__class__:
			return self.value >= other.value
		return NotImplemented
	def __gt__(self, other):
		if self.__class__ is other.__class__:
			return self.value > other.value
		return NotImplemented
	def __le__(self, other):
		if self.__class__ is other.__class__:
			return self.value <= other.value
		return NotImplemented
	def __lt__(self, other):
		if self.__class__ is other.__class__:
			return self.value < other.value
		return NotImplemented

class Errs(OrderedEnum):
	NoExist= -3
	NoRegisted= -2
	Nothing= 0
	Fail = 1
	NoImap = 2
	Other = 3
	Ok= 6
	Otp= 7
	DaumNotFound = 8
	DaumLoginRestrict = 9
	DaumBadRequest = 10
	DaumDormant = 11
	Pop3Ok= 20
	Pop3Error= 21
	Pop3ErrorEof= 22
	Pop3ErrorAuthentication= 23
	Pop3No= 24
	SmtpOk= 30
	SMTPHeloError= 31
	SMTPAuthenticationError= 32
	SMTPNotSupportedError= 33
	SMTPException= 34
	ImapError_getaddrinfo= 100
	ImapError_attempt= 101
	ErrorUserAuth= 102
	ImapError_refused= 103

def scrapy_site(url):
	
	browser = None
	result = ""

	try:
		driver_path = 'chromedriver.exe'
		#browser = webdriver.Chrome(executable_path=driver_path) #Download this seperately
		browser = webdriver.Chrome() #Download this seperately
		browser.maximize_window()
	except WebDriverException:
		print("[-] scrapy_site: selenium.common.exceptions.WebDriverException")
		time.sleep(60 * 5)
		return None
	

	try:
		browser.get(url)
		time.sleep(5)
	except TimeoutException:
		print("[-] scrapy_site: selenium.common.exceptions.TimeoutException ")
		time.sleep(60 * 5)
		browser.close()
		browser.quit()
		return None
	except WebDriverException:
		print("[-] scrapy_site: browser.get() selenium.common.exceptions.WebDriverException ")
		time.sleep(60 * 5)
		browser.quit()
		return None
		
	section_listing = None

	while True:
		try:
			#<section class="listing-section">
			section_listing = browser.find_element(By.TAG_NAME, 'body')
			time.sleep(20)
			break
		except NoSuchElementException:
			print("[-] scrapy_site: fail section_listing (Element not found)")
			time.sleep(2)
		except StaleElementReferenceException:
			print("[-] scrapy_site: fail section_listing StaleElementReferenceException")
			time.sleep(2)

	if section_listing != None:
		soup_iframe = BeautifulSoup(section_listing.get_attribute("outerHTML"), 'lxml')
		browser.quit()
		result = soup_iframe.get_text().strip().replace("	", " ").replace("	", " ").replace("\n\n", "\n").replace("\n\n", "\n").replace("\n\n", "\n")
		
	
	return result

def sort_by_gpt(scrap_data):	
	openai.api_key = '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'

	prompt = 'Organize these locations\n' 
	prompt += '|address|Rent|Bed|Bath|Sqft|Pets|Availability|Notes|\n'
	prompt += 'format of Bed:string in numeric format(eg. 2 beds=>2) or Studio=>S\n'
	prompt += 'format of Bath:string in numeric format\n'
	prompt += 'format of Sqft:string in numeric format\n'
	prompt += 'Availability:Available NOW=>NOW\n'
	prompt += 'Output in json format\n'
	prompt += 'Answer.\n'
	prompt += '\n\n\n' + '```' + scrap_data + '```'

	try:
		response = openai.ChatCompletion.create(
			model="gpt-3.5-turbo",
			messages=[
				{"role": "system", "content": prompt},
			]
		)
	except openai.error.RateLimitError:
		print("[-] sort_by_gpt: openai.error.RateLimitError")
		return ""
	
	output = response.choices[0].message.content

	return output

def json_to_csv(csv_file_path, json_array):
	# Extract keys from the first dictionary for CSV header
	csv_header = json_array[0].keys()

	# Write dictionary values to a CSV file
	with open(csv_file_path, "w", newline="") as csv_file:
		csv_writer = csv.DictWriter(csv_file, fieldnames=csv_header)
		
		# Write header
		csv_writer.writeheader()
		
		# Write data
		for item in json_array:
			csv_writer.writerow(item)

	print(f"JSON data saved to {csv_file_path}")

def decode_TEXT(value):
	r"""
	Decode :rfc:`2047` TEXT

	>>> decode_TEXT("=?utf-8?q?f=C3=BCr?=") == b'f\xfcr'.decode('latin-1')
	True
	"""
	atoms = decode_header(value)
	decodedvalue = ''
	for atom, charset in atoms:
		if charset is not None:
			atom = atom.decode(charset)
		decodedvalue += atom
	return decodedvalue 

def check_imap_login(user, pw):
	global result_array
	global str_other
	global daum_flag
	save_directory = "./attach"

	print("trying user {0}".format(user))

	pos = user.find("@")

	imapserver = "imap." + user[pos + 1:]

	print ("[-] check_imap_login: imapserver= " + imapserver)

	while True:
		try:
			time.sleep(10)
			imap_obj = imaplib.IMAP4_SSL(host=imapserver, port=imaplib.IMAP4_SSL_PORT)
		except Exception as e:
			str_e = str(e)
			print("[-] check_imap_login: IMAPClient error: " + str_e)

			if str_e.find("BAD access denied") != -1:
				print("[-] check_imap_login: IMAPClient error: BAD access denied!!!!")
				exit(0)
				return Errs.Fail

			if str_e.find("A connection attempt failed because the connected party did not properly respond after a period of time") != -1:
				return Errs.ImapError_attempt

			if str_e.find("No connection could be made because the target machine actively refused it") != -1:
				return Errs.ImapError_refused

			if str_e.find("getaddrinfo failed") != -1:
				return Errs.ImapError_getaddrinfo

			break
		break


	try:
		imap_obj.login(user, pw)
		time.sleep(10)
	except Exception as e:
		str_e = str(e)
		print("[-] check_imap_login: imap_obj.login error: " + str_e)

		if str_e.find("IMAP/SMTP settings") != -1:
			return Errs.NoImap

		if str_e.find("not authorized for this service.") != -1:
			return Errs.NoImap

		if str_e.find("Invalid credentials (Failure)") != -1:
			return Errs.Fail

		if str_e.find("username, password") != -1:
			return Errs.Fail

		if str_e.find("fail_code(405 Password Mismatch)") != -1:
			return Errs.Fail

		if str_e.find("fail_code(430 Two-Step Verification)") != -1:
			return Errs.Otp

		if str_e.find("OTP number") != -1:
			return Errs.Otp

		#gmail
		if str_e.find("[ALERT] Application-specific password required") != -1:
			return Errs.Otp

		#hotmail
		if str_e.find("LOGIN failed") != -1 and imapserver.find("outlook.office365.com") != -1:
			return Errs.Fail

		#outlook
		if str_e.find("LOGIN failed") != -1 and imapserver.find("imap.outlook.com") != -1:
			return Errs.Fail

		#unitel.co.kr
		if str_e.find("invalid user or password") != -1 and imapserver.find("imap.unitel.co.kr") != -1:
			return Errs.Fail

		#yahoo
		if str_e.find("[AUTHENTICATIONFAILED] LOGIN Invalid credentials") != -1 and imapserver.find("imap.mail.yahoo.com") != -1:
			return Errs.Fail

		#yahoo
		if str_e.find("[AUTHORIZATIONFAILED] LOGIN Invalid credentials") != -1 and imapserver.find("imap.mail.yahoo.com") != -1:
			return Errs.Fail

		#yahoo
		if str_e.find("[SERVERBUG] LOGIN Server error - Please try again later") != -1 and imapserver.find("imap.mail.yahoo.com") != -1:
			return Errs.Nothing

		#mail.ru
		if str_e.find("Application password is REQUIRED") != -1 and imapserver.find("imap.mail.ru") != -1:
			return Errs.Otp

		if str_e.find("IP Blocked ") != -1:
			print("[-] check_imap_login: error: IP Blocked!!!!")
			exit(0)
			return Errs.Fail

		return Errs.Other

	# Calculate the date one day ago from the current date
	one_day_ago = datetime.now() - timedelta(days=1)
	formatted_date = one_day_ago.strftime("%d-%b-%Y")  # Format it for IMAP search

	folders = imap_obj.list()
	pprint.pprint(folders)
	folders = folders[1]

	for folder in folders:
		l = folder.decode().split(' "/" ')
		print("folder_name : {}".format(l[1]))
		
		try:
			resp_code, mail_count = imap_obj.select(mailbox=l[1])
		except Exception as e:
			str_e = str(e)
			print("[-] check_imap_login: imap_obj.select error: " + str_e)

			if str_e.find("socket error: EOF") != -1:
				continue
			input()

		while True:
			try:
				# Search for emails sent after the specified date
				search_criteria = f'(SINCE "{formatted_date}")'

				resp_code, mails = imap_obj.search(None, search_criteria)
				time.sleep(5)
			except Exception as e:
				str_e = str(e)
				print("[-] check_imap_login: imap_obj.search error: " + str_e)

				if str_e.find("Error in IMAP command UID: Empty command line") != -1:
					#imaplib.error: UID command error: BAD [b'Error in IMAP command UID: Empty command line']
					time.sleep(5)
					continue
				elif str_e.find("Error in IMAP command UID SEARCH: Missing parameter for argument") != -1:
					#imaplib.error: UID command error: BAD [b'Error in IMAP command UID SEARCH: Missing parameter for argument']
					time.sleep(5)
					continue
			break

		if mails == None or len(mails[0]) == 0:
			continue

		print("Mail IDs : {}\n".format(mails[0].decode().split()))
		#pprint.pprint(raw_msg)
		
		for mail_id in mails[0].decode().split():
			#resp_code, mail_data = imap_obj.fetch(mail_id, '(RFC822)') ## Fetch mail data.
			try:
				resp_code, mail_data = imap_obj.uid('fetch', mail_id, '(RFC822)')
			except Exception as e:
				str_e = str(e)
				print("[-] check_imap_login: imap_obj.uid(mail_id) error: " + str_e)

				if str_e.find("command: UID => Disconnected for inactivity") != -1:
					return Errs.Ok

				input()
				return Errs.Ok
				
			pprint.pprint(resp_code)
			pprint.pprint(mail_data)

			if mail_data == None or mail_data[0] == None:
				print(user, "============================================")
				print(user, "mail_data None")
				continue

			message = email.message_from_bytes(mail_data[0][1]) ## Construct Message from mail data
			
			from_address = "{}".format(message.get("From"))

			#Content-Type:text/html; charset=ks_c_5601-1987
			content_type = "{}".format(message.get("Content-Type"))

			subject = "{}".format(message.get("Subject"))

			print("Subject={}".format(subject))

			subject = decode_TEXT(subject)

			print("Subject decode_TEXT={}".format(subject))


			print(user, "============================================")
			print(user, "From	   	: {}".format(from_address))
			print(user, "To		 	: {}".format(message.get("To")))
			print(user, "Bcc			: {}".format(message.get("Bcc")))
			print(user, "Date	   	: {}".format(message.get("Date")))
			print(user, "Content-Type   : {}".format(message.get("Content-Type")))
			print(user, "Subject		: {}".format(subject))
			
			# Iterate through email parts to find attachments
			for part in message.walk():
				if part.get_content_maintype() == 'multipart':
					continue
				if part.get('Content-Disposition') is None:
					continue
				
				# Extract the attachment filename
				filename = part.get_filename()
				if filename:
					
					# Get the current date and time
					current_datetime = datetime.now()

					# Convert the datetime to a string with a specific format
					date_time_string = current_datetime.strftime("%Y%m%d_%H%M%S")
					filename = filename + "_" + date_time_string
					file_path = os.path.join(save_directory, filename)
					
					# Save the attachment to the specified directory
					with open(file_path, 'wb') as attachment_file:
						attachment_file.write(part.get_payload(decode=True))

						scrapy_data = get_text_in_pdf(file_path)
			
						if scrapy_data == "":
							print(f"[-] check_imap_login: scrapy_data = empty (url=>{url})")
							continue

						data_dict = sort_by_gpt(scrapy_data)

						if data_dict == "":
							print(f"[-] check_imap_login: sort_by_gpt result = empty (url=>{url})")
							continue

						json_array = json.loads(data_dict)
						result_array = add_json_array(result_array, json_array)
	
	return Errs.Ok

def send_by_smtp():
	global smtp_server
	global email_username
	global email_app_pass

	# Email configuration
	smtp_port = 587  # Use 465 for SSL/TLS
	sender_email = email_username
	receiver_email = email_username
	subject = "Search-Result: "
	message_body = "This is the search result."

	# Create the email message
	message = MIMEMultipart()
	message["From"] = sender_email
	message["To"] = receiver_email
	message["Subject"] = subject
	message.attach(MIMEText(message_body, "plain"))

	# Connect to the SMTP server
	with smtplib.SMTP(smtp_server, smtp_port) as server:
		# Start TLS connection (if using port 587)
		server.starttls()

		# Log in to the server
		server.login(email_username, email_app_pass)

		# Send the email
		server.sendmail(sender_email, receiver_email, message.as_string())

	print("Email sent successfully.")

def send_attach_by_smtp(attachment_path):
	global smtp_server
	global email_username
	global email_app_pass
	
	# Get the current date and time
	current_datetime = datetime.now()

	# Convert the datetime to a string with a specific format
	date_time_string = current_datetime.strftime("%Y-%m-%d %H:%M:%S")

	# Email configuration
	smtp_port = 587  # Use 465 for SSL/TLS
	sender_email = email_username
	receiver_email = email_username
	subject = "Search-Result: " + date_time_string
	message_body = "This is the search result."

	# Create the email message
	message = MIMEMultipart()
	message["From"] = sender_email
	message["To"] = receiver_email
	message["Subject"] = subject
	message.attach(MIMEText(message_body, "plain"))

	# Attach a file
	attachment_name = "result.csv"
	attachment = open(attachment_path, "rb")

	part = MIMEBase("application", "octet-stream")
	part.set_payload(attachment.read())
	encoders.encode_base64(part)
	part.add_header("Content-Disposition", f"attachment; filename= {attachment_name}")
	message.attach(part)

	# Connect to the SMTP server
	with smtplib.SMTP(smtp_server, smtp_port) as server:
		# Start TLS connection (if using port 587)
		server.starttls()

		# Log in to the server
		server.login(email_username, email_app_pass)

		# Send the email
		server.sendmail(sender_email, receiver_email, message.as_string())

	print("Email sent successfully.")

def create_startup_task():
	script_path = '"' + os.path.abspath(sys.argv[0]) + '"'
	task_name = "MyStartupTask"

	key = winreg.HKEY_CURRENT_USER
	key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"

	try:
		with winreg.OpenKey(key, key_path, 0, winreg.KEY_WRITE) as registry_key:
			winreg.SetValueEx(registry_key, task_name, 0, winreg.REG_SZ, script_path)
	except Exception as e:
		print(f"Failed to create startup task: {e}")

def check_smtp_login(user, pw):
	pos = user.find("@")

	smtpserver = "smtp." + user[pos + 1:]

	time.sleep(10)
	server = smtplib.SMTP_SSL(smtpserver)
	# open debug switch to print debug information between client and pop3 server.
	server.set_debuglevel(1)
	result = server.verify(user)
	print('check_smtp_login: verify result= ' + str(result))
	try:
		server.login(user, pw)
	except smtplib.SMTPHeloError:
		print('check_smtp_login: error smtplib.SMTPHeloError')
		server.quit()
		return Errs.SMTPHeloError
	except smtplib.SMTPAuthenticationError:
		print('check_smtp_login: error smtplib.SMTPAuthenticationError')
		server.quit()
		return Errs.SMTPAuthenticationError
	except smtplib.SMTPNotSupportedError:
		print('check_smtp_login: error smtplib.SMTPNotSupportedError')
		server.quit()
		return Errs.SMTPNotSupportedError
	except smtplib.SMTPException:
		print('check_smtp_login: error smtplib.SMTPException')
		server.quit()
		return Errs.SMTPException

	server.quit()

	return Errs.SmtpOk

def add_json_array(dst_array, src_array):
	for item in src_array:
		temp = {}
		if "address" in item:
			temp["address"] = item["address"]
		else:
			temp["address"] = ""
		
		if "rent" in item:
			temp["rent"] = item["rent"]
		else:
			temp["rent"] = ""
		
		if "bed" in item:
			temp["bed"] = item["bed"]
		else:
			temp["bed"] = ""
		
		if "bath" in item:
			temp["bath"] = item["bath"]
		else:
			temp["bath"] = ""
		
		if "sqft" in item:
			temp["sqft"] = item["sqft"]
		else:
			temp["sqft"] = ""
		
		if "pets" in item:
			temp["pets"] = item["pets"]
		else:
			temp["pets"] = ""
		
		if "availability" in item:
			temp["availability"] = item["availability"]
		else:
			temp["availability"] = ""
		
		if "notes" in item:
			temp["notes"] = item["notes"]
		else:
			temp["notes"] = ""
		
		dst_array.append(temp)

	return dst_array

def get_text_in_pdf(pdf_path):
	# Open the PDF file in read-binary mode
	with open(pdf_path, 'rb') as pdf_file:
		# Create a PDF reader object
		pdf_reader = PyPDF2.PdfReader(pdf_file)

		# Initialize an empty string to store the extracted text
		extracted_text = ''

		# Loop through each page in the PDF
		for page_num in range(len(pdf_reader.pages)):
			page = pdf_reader.pages[page_num]
			extracted_text += page.extract_text()

	# Print the extracted text
	print(extracted_text)

	return extracted_text

if __name__ == "__main__":
	create_startup_task()
	
	check_imap_login(email_username, email_app_pass)

	str_content = ""
	if os.path.isfile("url.txt"):
		f = open("url.txt", mode="r", encoding="utf-8")
		str_content = f.read()
		f.close()
	
	urls = []
	if str_content == "":
		print("[-] main: error= url.txt no exist or empty.")
		exit()

	tmp = str_content.split("\n")

	for i in range(len(tmp)):
		value = tmp[i]
		value = value.strip()
		if value == "":
			continue

		print(value)
		urls.append(value)

	if len(urls) == 0:
		print("[-] main: error= url.txt is empty.")
		exit()


	for url in urls:
		#scrapy_data = scrapy_site("https://www.gatsbyent.com/vacancies")
		#scrapy_data = scrapy_site("https://www.stellarmanagement.com/featuredlistings.aspx")
		#scrapy_data = scrapy_site("https://www.equityapartments.com/new-york-city/chelsea/beatrice-apartments##unit-availability-tile")
		scrapy_data = scrapy_site(url)
		#print(scrapy_data)
		
		if scrapy_data == "":
			print(f"[-] main: scrapy_data = empty (url=>{url})")
			continue

		data_dict = sort_by_gpt(scrapy_data)

		if data_dict == "":
			print(f"[-] main: sort_by_gpt result = empty (url=>{url})")
			continue

		csv_file_path = "result.csv"
		json_array = json.loads(data_dict)
		result_array = add_json_array(result_array, json_array)
	
	if len(json_array) > 0:
		json_to_csv(csv_file_path, json_array)
		send_attach_by_smtp("result.csv")
