
"""
Created on Tue Feb 18 10:53:56 2025

@author: OLLESC
"""

"""
Fixing kernel issues through the CMD. Basically take the python.exe pathway and add the following: -m pip install spyder-kernels==3.1.*

C:/Users/olle_/AppData/Local/Programs/Python/Python311/python.exe -m pip install spyder-kernels==3.1.*

Thereafter installing non-default libraries.

python -m pip install pyperclip PyPDF2 python-dateutil openpyxl watchdog pandas pyautogui

Also possibly uninstall all Python programs from Installed Apps and install 3.11 from the link below.
python-3.11.9-embed-amd64.zip
https://www.python.org/ftp/python/3.11.9/

Check the system variables. Might need to add the following into the system variables. But remember to flip the dashes.
C:/Users/olle_/AppData/Local/Programs/Python/Python311/Scripts/
C:/Users/olle_/AppData/Local/Programs/Python/Python311/
"""

import re
import pyperclip
import glob
from PyPDF2 import PdfReader
import os
import math
import datetime as dt
from dateutil.relativedelta import relativedelta
import shutil
import openpyxl
import tkinter as tk
from tkinter import ttk
import threading
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import time
import pandas as pd
import pyautogui
import ctypes
from pathlib import Path
from uuid import UUID

"""
Make it possible to choose which PDF in downloads you want to use.
"""

" Global variables. "
current_user = os.getlogin() # Extracting the username from the windows system.
downloads_path = str(Path.home() / "Downloads")
borgens_sft_path = str(Path.home() / "Desktop" / "Borgens SFT BGF.xlsx")

# =============================================================================
# downloads_path = "C:/Users/" + current_user + "/Downloads" # Path to the Downloads folder.
# =============================================================================

" A function used for showing cursor notifications. It is being kept outside of classes. "

def show_cursor_notification(text):
	"""Shows a small notification near the cursor in system screen coordinates."""
	# Get the cursor's position directly via pyautogui (system coordinates)
	x, y = pyautogui.position()
	# Create a new Toplevel window for the notification
	notification = tk.Toplevel()
	notification.overrideredirect(True) # Remove window borders
	notification.wm_attributes("-topmost", 1) # Ensure it stays on top
	# Get window handle and apply WS_EX_LAYERED | WS_EX_TRANSPARENT
	notification.update_idletasks()
	hwnd = ctypes.windll.user32.GetParent(notification.winfo_id())
	extended_style = 0x00000020 | 0x00000080 # WS_EX_LAYERED | WS_EX_TRANSPARENT
	ctypes.windll.user32.SetWindowLongPtrW(hwnd, -20, extended_style)
	notification.geometry(f"+{x}+{y - 80}")
	# Create the label inside the Toplevel window
	label = tk.Label(notification, text=text, font=("Arial", 12, "italic"), fg="black", bg="white", 
				  relief="solid", padx=6, pady=4)
	label.pack()
	# Make the notification disappear after 2 seconds.
	notification.after(4000, notification.destroy)
	
" A function for extracting the desktop path including the onedrive segment of the pathway, appending Borgen Temp Saver and ensuring the folder exists. "

def get_sharepoint_borgen_folder_path():

	from datetime import datetime
	from getpass import getuser
	from pathlib import Path
	
	shortname = getuser()
	this_year = str(datetime.now().year)
	
	sharepoint_path = Path("C:/Users") / shortname / "Circle K Europe/SE - Krav och kredit"
	target_path = sharepoint_path / f"Borgen {this_year} Riga"

	# Ensure the folder exists
	target_path.mkdir(parents=True, exist_ok=True)

	return target_path

# Storing the PDF Saver Path.
pdf_saver_path_sharepoint = get_sharepoint_borgen_folder_path()

def get_desktop_borgen_folder_path():

	from datetime import datetime
	from getpass import getuser
	from pathlib import Path
	
	shortname = getuser()
	this_year = str(datetime.now().year)
	
	desktop_path = Path("C:/Users") / shortname / "Desktop"
	target_path = desktop_path / f"Desktop Borgens {this_year} Riga"

	# Ensure the folder exists
	target_path.mkdir(parents=True, exist_ok=True)

	return target_path

# Storing the PDF Saver Path.
pdf_saver_path_desktop = get_desktop_borgen_folder_path()

" A function for extracting the desktop path including the onedrive segment of the pathway, checking for the Excel file 'Borgens SFT BGF.xlsx', and creating it if it does not exist. "

'https://acteurope.sharepoint.com/:x:/r/sites/CKE-EU-FCAC-Credit-Department/Shared%20Documents/Credit%20Riga%20Team/SCA%20team/SE%20team/SWE%20Follow%20up.xlsx?d=w5c364fb6f2fa455ab380c7781c920bda&csf=1&web=1&e=e8rTnt'

def get_excel_file_path():
	# Desktop folder GUID
	FOLDERID_Desktop = '{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}'

	# Convert GUID string to GUID structure
	guid = UUID(FOLDERID_Desktop).bytes_le

	ptr = ctypes.c_wchar_p()
	result = ctypes.windll.shell32.SHGetKnownFolderPath(guid, 0, None, ctypes.byref(ptr))
	if result != 0:
		raise OSError("Could not get Desktop folder path")

	desktop_path = Path(ptr.value)
	ctypes.windll.ole32.CoTaskMemFree(ptr)

	# Define the Excel file path directly on the Desktop
	excel_file_path = desktop_path / "Borgens SFT BGF.xlsx"

	# Check if the Excel file exists
	if not excel_file_path.exists():
		# If not, create a new Excel file
		# Create a new workbook
		workbook = openpyxl.Workbook()
		sheet = workbook.active
		# Set the header row
		headers = ["ORG ID", "Follow-up", "Completion", "Person"]
		sheet.append(headers)

		# Save the workbook to the Desktop
		workbook.save(excel_file_path)
		print(f"Excel file created: {excel_file_path}")
	else:
		print(f"Excel file already exists: {excel_file_path}")

	return excel_file_path

# Example usage:
excel_file_path = get_excel_file_path()

def bring_cloud_excel_to_desktop():
	# Fixed Desktop path
	desktop = Path(f"C:/Users/{current_user}/Desktop/")
	
	# Detect OneDrive path
	one_drive = Path.home() / "OneDrive"
	if not one_drive.exists():
		one_drive = Path.home() / "OneDrive - YourCompany"
	if not one_drive.exists():
		raise OSError("OneDrive folder not found.")
	
	# Cloud file path
	cloud_file = one_drive / "SWE Follow Up.xlsx"
	
	# Desktop file path
	desktop_file = desktop / "SWE Follow Up.xlsx"
	
	# If the cloud file exists, copy it to Desktop
	if cloud_file.exists():
		shutil.copy(cloud_file, desktop_file)
		print(f"File copied to Desktop: {desktop_file}")
		return desktop_file
	
	# If cloud file does not exist, create a fresh one on Desktop
	workbook = openpyxl.Workbook()
	sheet = workbook.active
	headers = ["ORG ID", "Follow-up", "Completion", "Person"]
	sheet.append(headers)
	workbook.save(desktop_file)
	
	print(f"Cloud file not found. New Excel file created on Desktop: {desktop_file}")
	return desktop_file

bring_cloud_excel_to_desktop()

" Part 1. The class for storing and returning a dictionary of the information in UC. "

class UCDataExtractor:
	def __init__(self):
		self.scraper_data_company = None
		self.scraper_data_borgensman = None
		self.update_data() # Automatically fetch clipboard data on initialization.

	def update_data(self, scraper_data=None, selected_row_str=None):
		"""Updates the class with new data from the provided argument or clipboard if not provided."""
		if selected_row_str:
			# If selected_row_str is passed, split it accordingly
			self.scraper_data_company, self.scraper_data_borgensman = selected_row_str.split(" ||| ")
		else:
			# If no selected_row_str, use the provided scraper_data or fetch from clipboard
			if scraper_data is None:
				scraper_data = pyperclip.paste()
			if ' ||| ' in scraper_data:
				scraper_data = scraper_data.strip('"')
			else:
				# Default placeholder if the data format is not recognized
				scraper_data = "No Company Copied AB | Org.nummer: 999999-9999 ||| None Copied | Personnummer 888888-8888"
			
			# Split the data into company and borgensman parts
			self.scraper_data_company, self.scraper_data_borgensman = scraper_data.split(" ||| ")
			
		return self.scraper_data_company, self.scraper_data_borgensman
	
	def UC_Data_Dictionary(self, selected_row_str=None):
		"""Parses UC data into a dictionary, removing spaces near dashes."""
		# If selected_row_str is passed, update the data first
		if selected_row_str:
			self.update_data(selected_row_str=selected_row_str)
		
		""" Parses UC data into a dictionary, removing spaces near dashes. """
		uc_org_match = re.search(r'Org\.nummer:\s*(.*)', self.scraper_data_company)
		uc_org_nr = uc_org_match.group(1) if uc_org_match else None
		uc_original_org_nr = uc_org_nr # Keeping the original format

		uc_company_name_match = re.search(r'^(.*?)\s*\|', self.scraper_data_company)
		uc_company_name = uc_company_name_match.group(1) if uc_company_name_match else None
		uc_original_cn = uc_company_name # Keeping the original format

		uc_pn_match = re.search(r'Personnummer\s*(.*)', self.scraper_data_borgensman)
		uc_personnummer = uc_pn_match.group(1) if uc_pn_match else None

		uc_borgensman_name_match = re.search(r'^(.*?)\s*\|', self.scraper_data_borgensman)
		uc_borgensman_personal_name = uc_borgensman_name_match.group(1) if uc_borgensman_name_match else None

		# Adaptation for tolerating a space between a dash.
		if uc_company_name:
			uc_company_name = uc_company_name.replace(" -", "-").replace("- ", "-")
		if uc_borgensman_personal_name:
			uc_borgensman_personal_name = uc_borgensman_personal_name.replace(" -", "-").replace("- ", "-")

		return {
			"uc_org_match": uc_org_match,
			"uc_company_name_match": uc_company_name_match,
			"uc_pn_match": uc_pn_match,
			"uc_borgensman_name_match": uc_borgensman_name_match,
			"uc_org_nr": uc_org_nr,
			"uc_original_org_nr": uc_original_org_nr,
			"uc_company_name": uc_company_name,
			"uc_original_cn": uc_original_cn,
			"uc_personnummer": uc_personnummer,
			"uc_borgensman_personal_name": uc_borgensman_personal_name,
		}

" Part 2. The class for storing and returning a dictionary of the PDF. "

class PDFDataExtractor:
	
	def __init__(self):
		self.data = {}
	
	def newest_file_extractor(self, chosen_path=None):
		"""
		This function extracts the newest file (by last modified date) within the 
		directory 'downloads_path' and returns the full path to that file.
		The file must contain the word "Borgensförbindelse" in its name to be considered.
		If a near-identical file exists, it deletes the conflicting PDF before renaming.
		"""
		if chosen_path is not None:
			return chosen_path  # Return provided path directly

		files = glob.glob(os.path.join(downloads_path, '*'))
		matching_files = [file for file in files if "Borgensförbindelse" in os.path.basename(file)]

		if not matching_files:
			return None

		# Get the newest file based on last modified date
		newest_file = max(matching_files, key=os.path.getmtime)
		base_name_newest = os.path.splitext(os.path.basename(newest_file))[0]
		new_file_path = os.path.join(downloads_path, base_name_newest + ".pdf")

		# Look for a near-identical file (same base name, any extension except the newest one)
		for file in matching_files:
			if file == newest_file:
				continue
			if os.path.splitext(os.path.basename(file))[0] == base_name_newest:
				if file.endswith(".pdf"):
					os.remove(file)
					print(f"Deleted conflicting PDF: {file}")
				else:
					previous_file = file + " Previous File"
					os.rename(file, previous_file)
					print(f"Renamed previous file: {previous_file}")

		# Rename newest file to .pdf
		if newest_file != new_file_path:
			os.rename(newest_file, new_file_path)
			print(f"Renamed newest file to: {new_file_path}")

		return new_file_path if os.path.exists(new_file_path) else None

	def PDF_Extractor(self, chosen_path = None):
		"""
		This function extracts text from the latest PDF file found in 'downloads_path'.
		It returns a single string with the text contents of the file.
		"""

		if chosen_path == None:
			newest_file_path = self.newest_file_extractor(chosen_path)
			if not newest_file_path:
				return None
			
			reader = PdfReader(newest_file_path)
			
			all_text = ''
			for page in reader.pages:
				extracted_text = page.extract_text()
				if extracted_text:
					all_text += extracted_text
		else:
			reader = PdfReader(chosen_path)
			all_text = ''
			for page in reader.pages:
				extracted_text = page.extract_text()
				if extracted_text:
					all_text += extracted_text
		
		all_text = all_text.replace('\x00','fi') # Replacing the special character with 'fi'.
		
		" Removing any \n for the signature name if the name is too long. "
		all_text = re.sub(r"(1\..*?)\n+(?=.*?Signed)", lambda m: m.group(1).replace('\n', ' '), all_text, flags=re.DOTALL)
		
		return all_text
	
	def personal_number_standardizer(self, personnummer):
		personnummer = str(personnummer)
		"""
		This function takes a Swedish personal number (personnummer) as input and 
		ensures that it follows a standard format: YYMMDD-ABCD.
		It removes any country code if present and adds a hyphen if needed.
		"""
		# Changing a space to a dash for orgnrs with a space.
		if (len(personnummer) == 11 and personnummer[-5] == " " and personnummer[:2] == "55"	and personnummer.replace(" ", "").isdigit()):  # all other chars are digits
			personnummer = personnummer[:6] + '-' + personnummer[7:]
		if personnummer:
			# If the length is not within this 10 to 13, do not standardize.
			if len(personnummer) not in range(10,14):
				return (personnummer, False)
			# If dash the numbers before the dash are not 6 or 8, do not standardize.
			elif ('-' in personnummer) and len(personnummer.split('-')[0]) not in (6,8):
				return (personnummer, False)
			# If dash and the numbers after the dash are not 4, do not standardize.
			elif ('-' in personnummer) and len(personnummer.split('-')[1]) != 4:
				return (personnummer, False)
			# If space the numbers before the space are not 6 or 8, do not standardize.
			elif (' ' in personnummer) and len(personnummer.split(' ')[0]) not in (6,8):
				return (personnummer, False)
			# If space and the numbers after the space are not 4, do not standardize.
			elif (' ' in personnummer) and len(personnummer.split(' ')[1]) != 4:
				return (personnummer, False)
			# If no dash and no space and the length is 11, do not standardize.
			elif ('-' not in personnummer and ' ' not in personnummer) and len(personnummer) == 11:
				return (personnummer, False)
			# If no dash and no space and the length is 13, do not standardize.
			elif ('-' not in personnummer and ' ' not in personnummer) and len(personnummer) == 13:
				return (personnummer, False)
			# If the personnummer has a dash is in an incorrect location, do not standardize.
			elif ("-" in personnummer and personnummer[-5] != "-"):
				return (personnummer, False)
			# If the personnummer has a space is in an incorrect location, do not standardize.
			elif (" " in personnummer and personnummer[-5] != " "):
				return (personnummer, False)
			# If the personal ID number is long and the person is not of legal age, do not standardize.
			elif len(personnummer) in range(12, 14) and int(personnummer[2:4]) > dt.datetime.now().year - 18:
				return (personnummer, False)
			# If the personal ID number is short and the person is not of legal age, do not standardize.
			elif len(personnummer) in range(10, 12) and int(personnummer[0:2]) > dt.datetime.now().year - 18:
				return (personnummer, False)
			# If the personal ID number is long, and the person is born in the 1900s and is too old, do not standardize.
			elif len(personnummer) in range(12, 14) and personnummer[:2].startswith("19") and int(personnummer[:4][2:4]) < (dt.datetime.now().year - 18) % 100:  # 1900–1907 as of year 2025 → impossible:
				return (personnummer, False)
			# If the format is short YYMMDDABCD with 10 digits and no dash, then standardize.
			elif (len(personnummer) == 10 and personnummer[-5] != "-"): 
				if int(personnummer[2:4]) > 12:
					return (personnummer, False)
				elif int(personnummer[4:6]) > 31:
					return (personnummer, False)
				else:
					return (personnummer[:-4] + "-" + personnummer[-4:], True)
			# If the format is short YYMMDD-ABCD with 10 digits and a dash, then standardize.
			elif (len(personnummer) == 11 and personnummer[-5] == "-"):
					if int(personnummer[2:4]) > 12:
						return (personnummer, False)
					elif int(personnummer[4:6]) > 31:
						return (personnummer, False)
					else:
						return (personnummer, True)
			# If the format is short YYMMDD ABCD with 10 digits and a space, then standardize.
			elif (len(personnummer) == 11 and personnummer[-5] == " "):
					if int(personnummer[2:4]) > 12:
						return (personnummer, False)
					elif int(personnummer[4:6]) > 31:
						return (personnummer, False)
					else:
						return (personnummer[:6] + '-' + personnummer[7:], True)
			elif len(personnummer) in range(12,14):
				# If the personal ID number is long and the person is not born in the 1900s or 2000s, then do not standardize the personal ID number.
				if ((len(personnummer) == 12 or (len(personnummer) == 13) and personnummer[-5] == "-")) and personnummer[:2] not in ('19', '20'):
					return (personnummer, False)
				# If the format is long YYYYMMDDABCD with 12 digits and no dash, then standardize.
				elif (len(personnummer) == 12 and personnummer[-5] != "-"):
						if int(personnummer[4:6]) > 12:
							return (personnummer, False)
						elif int(personnummer[6:8]) > 31:
							return (personnummer, False)
						else:
							return (personnummer[2:][:-4] + "-" + personnummer[2:][-4:], True) # Excluding the first two digits signifying the century and adding a dash in the fifth spot from the last.
				# If the format is long YYYYMMDD-ABCD with 12 digits and a dash, then standardize.
				elif (len(personnummer) == 13 and personnummer[-5] == "-"):
						if int(personnummer[4:6]) > 12:
							return (personnummer, False)
						elif int(personnummer[6:8]) > 31:
							return (personnummer, False)
						else:
							return (personnummer[2:], True) # Excluding the first two digits signifying the century.
				# If the format is long YYYYMMDD ABCD with 12 digits and a space, then do not standardize.
				elif (len(personnummer) == 13 and personnummer[-5] == " "):
						if int(personnummer[4:6]) > 12:
							return (personnummer, False)
						elif int(personnummer[6:8]) > 31:
							return (personnummer, False)
						else:
							return (personnummer[2:8] + '-' + personnummer[9:], True) # Excluding the first two digits signifying the century.
				else:
					return (personnummer, False)
		
	def PDF_Data_Dictionary(self, chosen_path = None):
		"""
		This function extracts relevant fields from a PDF containing company data.
		It uses regex patterns to find information such as organization number, 
		personal number, names, and signatures, then stores them in a dictionary.
		"""
		all_text_pdf = self.PDF_Extractor(chosen_path)
		if all_text_pdf:
			all_text_pdf = all_text_pdf.replace("\r", "").replace("\n", "") # Remove new line spacing.

		if not all_text_pdf:
			# print('No PDF in Downloads')

			self.data = {
				'pdf_org_nr': 'No PDF in Downloads',
				'pdf_personal_number': 'No PDF in Downloads',
				'pdf_personal_name': 'No PDF in Downloads',
				'pdf_company_name': 'No PDF in Downloads',
				'pdf_signature_name': 'No PDF in Downloads',
				'pdf_signature_dob': 'No PDF in Downloads',
				'pdf_personal_number_standardized': ('No PDF in Downloads', False),
				'pdf_original_org_nr': 'No PDF in Downloads',
				'pdf_original_personal_name': 'No PDF in Downloads',
				'pdf_original_cn': 'No PDF in Downloads',
				'pdf_original_personnummer_standardized': ('No PDF in Downloads', False),
				'pdf_original_personnummer': 'No PDF in Downloads',
				'pdf_company_name_upper': 'No PDF in Downloads'
			}
			return self.data
		"""
        I have added an if else statement here. The structure of the PDF Changed with the minor different in "METHOD DETAILS" 
        losing the space in between the words. 'SHA-512' is only present in the newer version of the PDF.
		"""

		if 'SHA-512' in self.PDF_Extractor():
			print('SHA-512')
			patterns = {
				"pdf_org_nr": r"Org\. nr:\s*(.*?)(?=\s*Företagsnamn)",
				"pdf_personal_number": r"Pers\. nr:\s*(.*?)(?=\s*Namn:)",
				"pdf_personal_name": r"Namn:\s*(.*?)(?=\s*(?:Adress|Tel|$))",
				"pdf_company_name": r"Företagsnamn:\s*(.*?)(?=\s*(?:Adress|Tel|$))",
				"pdf_signature_name": r"METHODDETAILS\s*1\.\s*(.*?)\s*Signed",
				"pdf_signature_dob": r"Swedish BankID \(DOB:\s*(\d{4}/\d{2}/\d{2})\)"
			}
		else:
			print('else')
			patterns = {
				"pdf_org_nr": r"Org\. nr:\s*(.*?)(?=\s*Företagsnamn)",
				"pdf_personal_number": r"Pers\. nr:\s*(.*?)(?=\s*Namn:)",
				"pdf_personal_name": r"Namn:\s*(.*?)(?=\s*(?:Adress|Tel|$))",
				"pdf_company_name": r"Företagsnamn:\s*(.*?)(?=\s*(?:Adress|Tel|$))",
				"pdf_signature_name": r"METHOD DETAILS\s*1\.\s*(.*?)\s*Signed",
				"pdf_signature_dob": r"Swedish BankID \(DOB:\s*(\d{4}/\d{2}/\d{2})\)"
			}
			
		for key, pattern in patterns.items():
			match = re.search(pattern, all_text_pdf, re.DOTALL)
			# match = re.search(pattern, all_text_pdf)
			self.data[key] = match.group(1) if match else None
		
		# Standardizing the personal ID number. It will only come out as standardized when it is not invalid.
		if self.data["pdf_personal_number"]:
			self.data["pdf_personal_number_standardized"] = self.personal_number_standardizer(self.data["pdf_personal_number"])
		
		# Removing maximally one space between capitalized and non capitalized letters.
		self.data["pdf_personal_name"] = re.sub(r'([A-Z]) ([a-z])', r'\1\2', self.data["pdf_personal_name"]) if self.data["pdf_personal_name"] else None
		self.data["pdf_company_name"] = re.sub(r'([A-Z]) ([a-z])', r'\1\2', self.data["pdf_company_name"]) if self.data["pdf_company_name"] else None

		# Storing original versions of the variables.
		self.data["pdf_original_org_nr"] = self.data["pdf_org_nr"]
		self.data["pdf_original_personal_name"] = self.data["pdf_personal_name"]
		self.data["pdf_original_cn"] = self.data["pdf_company_name"]
		self.data["pdf_original_personnummer_standardized"] = self.personal_number_standardizer(self.data["pdf_personal_number"])
		self.data["pdf_original_personnummer"] = self.data["pdf_personal_number"] # Ensure `pdf_personal_number_re` exists
		
		# Calling and storing the UC dictionary.
		UC_Extractor = UCDataExtractor()
		UC_Data_Dictionary = UC_Extractor.UC_Data_Dictionary()
		
		# If there are extra spaces in the company name from the PDF but the characters are identical to UC, then use the UC version.
		uc_company_name = (UC_Data_Dictionary["uc_company_name_match"].group(1) if UC_Data_Dictionary["uc_company_name_match"] else None)
		if uc_company_name and uc_company_name.upper() != self.data["pdf_company_name"].upper():
			if self.data["pdf_original_cn"].replace(' ', '').upper() == UC_Data_Dictionary["uc_company_name"].replace(' ', '').upper():
				if self.data["pdf_original_cn"].isupper():
					self.data["pdf_original_cn"] = uc_company_name.upper()
				else:
					self.data["pdf_original_cn"] = uc_company_name
		
		# Removing one space in the company name for visualization in the GUI.
		if self.data["pdf_original_cn"].upper() != UC_Data_Dictionary["uc_company_name"].upper():
			name_parts = UC_Data_Dictionary["uc_company_name"].split()
			if len(name_parts) > 1 and name_parts[1] == "".join(self.data["pdf_company_name"].split()[1:3]):
				self.data["pdf_original_cn"] = re.sub(
					r"(?<!\S)([bcdfghjklmnpqrstvwxyzBCDFGHJKLMNPQRSTVWXYZ])\s+([A-Za-zÅÄÖåäö])",
					r"\1\2",
					self.data["pdf_company_name"],
					count=1
				)
		
		# Adaptation for tolerating a spaces beside a dash in the company name.
		self.data["pdf_company_name"] = self.data["pdf_company_name"].replace(" -", "-").replace("- ", "-").replace(" -", "-").replace("- ", "-")
		# Capitalized version of the company name for verification.
		self.data["pdf_company_name_upper"] = self.data["pdf_company_name"].upper()
		# Adaptation for tolerating a spaces beside a dash in the company name.
		
		# Return final extracted data dictionary.
		return self.data

PDFDataExtractor().PDF_Extractor().replace("\r", "").replace("\n", "")
PDFDataExtractor().PDF_Data_Dictionary()

""" Part 3. A class for sensing the clipboard for the input of UC data. """

class ClipboardMonitor:
	def __init__(self, app_instance):
		self.app = app_instance # Store the reference to PersonalGuaranteeApp
		self.orgnr = None
		self.personnummer = None
		self.original_clipboard = pyperclip.paste() # Save original clipboard once
		self.previous_clipboard = self.original_clipboard # Use this after first run
		self.current_clipboard = None  
		self.monitoring_active = True
		self.first_reset_done = False # Track if original clipboard has been restored
		# Start monitoring in a separate thread
		self.extractor_running = threading.Event()
		self.monitoring_thread = threading.Thread(target=self.monitor_clipboard, daemon=True)
		self.monitoring_thread.start()

	def process_clipboard(self, current_clipboard):
		"""Processes clipboard text to extract Org.nummer or Personnummer."""
		# Check if it's an Org.nummer and reset Personnummer if necessary
		if ' |  Org.nummer: ' in current_clipboard:
			if "Skriv ut" in current_clipboard:
				match = re.search(r"Skriv ut(.*?\d{6}-\d{4})", current_clipboard, re.DOTALL)
			else:
				match = re.search(r"(.*?\d{6}-\d{4})", current_clipboard, re.DOTALL)
			if match:
				new_orgnr = match.group(1).replace('\t', "").replace('\r\n', '').replace('  |  ', ' | ')
				if new_orgnr != self.orgnr: # If a new org.nummer is found, reset personnummer
					self.orgnr = new_orgnr
					if not (self.orgnr and self.personnummer):
						show_cursor_notification('Org Nr Copied from UC.')
					print(f"New Org.nummer: {self.orgnr}")

		# Check if it's a Personnummer and reset Org.nummer if necessary
		if ' |  Personnummer ' in current_clipboard:
			if "Skriv ut" in current_clipboard:
				match = re.search(r"Skriv ut(.*?\d{6}-\d{4})", current_clipboard, re.DOTALL)
			else:
				match = re.search(r"(.*?\d{6}-\d{4})", current_clipboard, re.DOTALL)
			if match:
				new_personnummer = match.group(1).replace('\t', "").replace('\r\n', '').replace('  |  ', ' | ')
				if new_personnummer != self.personnummer: # If a new personnummer is found, reset orgnr
					self.personnummer = new_personnummer
					if not (self.orgnr and self.personnummer):
						show_cursor_notification('Personnummer Copied from UC.')
					print(f"New Personnummer: {self.personnummer}")

		# If both Org.nummer and Personnummer are found, combine and copy result
		if self.orgnr and self.personnummer:
			combined_result = f'"{self.orgnr} ||| {self.personnummer}"'
			print(f"Combined Result: {combined_result}")
			pyperclip.copy(combined_result)
			# Ensure update_gui is only called once during the current cycle
			self.app.update_gui()
			show_cursor_notification('Information Updated.')
			if not self.first_reset_done:
				pyperclip.copy(self.original_clipboard)
				self.first_reset_done = True # Set flag to ensure it happens only once
			else:
				pyperclip.copy(self.previous_clipboard)
			# Reset values for next detection
			self.orgnr = None
			self.personnummer = None

	def restore_clipboard(self):
		"""Restores clipboard to its original content only once, then always to previous_clipboard."""
		if not self.first_reset_done:
			pyperclip.copy(self.original_clipboard) # Restore only once
			self.first_reset_done = True # Mark that the original clipboard has been restored
			print("Clipboard restored to original content.")
		else:
			pyperclip.copy(self.previous_clipboard) # Use previous_clipboard thereafter
			print("Clipboard restored to previous content.")

	def monitor_clipboard(self):
		"""Monitors clipboard and processes relevant copied text."""
		while self.monitoring_active:
			time.sleep(0.2)	# Prevent excessive CPU usage
			
			current_clipboard = pyperclip.paste().strip()
	
			# Only process new clipboard contents
			if current_clipboard and current_clipboard != self.current_clipboard:
				self.current_clipboard = current_clipboard
	
				# Process clipboard if it's in the expected format
				if ' |  Org.nummer: ' in current_clipboard or ' |  Personnummer ' in current_clipboard:
					self.process_clipboard(current_clipboard)
				else:
					print("Unrelated text copied. Resetting detection...")
					self.previous_clipboard = current_clipboard

# =============================================================================
# 	def monitor_clipboard(self):
# 		"""Monitors clipboard and processes relevant copied text."""
# 		import keyboard
# 		while self.monitoring_active:
# 			time.sleep(0.1) # Prevent excessive CPU usage
# 			if keyboard.is_pressed('ctrl') and keyboard.is_pressed('c'):
# 				time.sleep(0.1) # Allow clipboard update
# 				current_clipboard = pyperclip.paste().strip()
# 
# 				# Only process new clipboard contents
# 				if current_clipboard and current_clipboard != self.current_clipboard:
# 					self.current_clipboard = current_clipboard  
# 
# 					# Process clipboard if it's in the expected format
# 					if ' |  Org.nummer: ' in current_clipboard or ' |  Personnummer ' in current_clipboard:
# 						self.process_clipboard(current_clipboard)
# 					else:
# 						print("Unrelated text copied. Resetting detection...")
# 						# print('The new unrelated clipboard is:', self.current_clipboard)
# 						# Update previous_clipboard only when unrelated text is copied
# 						self.previous_clipboard = current_clipboard  
# 						pyperclip.copy(self.previous_clipboard)
# =============================================================================

	def start_monitoring(self):
		"""Starts the clipboard monitoring process in a separate thread."""
		self.monitoring_thread = threading.Thread(target=self.monitor_clipboard, daemon=True)
		self.monitoring_thread.start()

	def stop_monitoring_clipboard(self):
		"""Stops the clipboard monitoring process."""
		if self.extractor_running.is_set():
			self.extractor_running.clear() # Clear the event flag to stop any waiting threads

		# Stop the monitoring thread by setting the active flag to False
		self.monitoring_active = False
		
		# Ensure the monitoring thread finishes before proceeding
		if self.monitoring_thread and self.monitoring_thread.is_alive():
			self.monitoring_thread.join(timeout=1) # Wait for the thread to finish
		print(self.orgnr)
		print(self.personnummer)
		print("Exiting process...")
		
		# Restore the clipboard to its previous state
		self.restore_clipboard()

	def run(self):
		"""Main loop to keep the script running."""
		print('Run Started.')
		self.start_monitoring()
		while self.monitoring_active:
			time.sleep(5) # Keeps main thread alive
		print("Clipboard monitoring has been stopped.")

""" Part 4. A class for verifying the data from UC corresponding with the data from the borgensförbindelse pdf. """

class Verifier:

	def __init__(self, chosen_path=None, selected_row_str=None):
		self.UC_Data_Dictionary = UCDataExtractor().UC_Data_Dictionary(selected_row_str)
		self.PDF_Data_Dictionary = PDFDataExtractor().PDF_Data_Dictionary(chosen_path)
 
	def org_nr_verifier(self):

		uc_org_nr = self.UC_Data_Dictionary["uc_org_nr"]
		pdf_org_nr = self.PDF_Data_Dictionary["pdf_org_nr"]

		if uc_org_nr is not None and pdf_org_nr is not None:
			return uc_org_nr == pdf_org_nr
		else:
			# Handle the case where one or both keys don't exist
			return False

	def company_name_verifier(self):

		uc_company_name = self.UC_Data_Dictionary["uc_company_name"]
		pdf_company_name = self.PDF_Data_Dictionary["pdf_company_name"]

		if uc_company_name and pdf_company_name:
			uc_company_name = uc_company_name.upper()
			pdf_company_name = pdf_company_name.upper()
		else:
			# Handle the case where one or both keys don't exist
			return False

		# Removing spaces between the dashes in the PDF and UC.
		pdf_company_name = pdf_company_name.replace(" -", "-")
		pdf_company_name = pdf_company_name.replace("- ", "-")
		uc_company_name = uc_company_name.replace(" -", "-")
		uc_company_name = uc_company_name.replace("- ", "-")

		# Replacing specific company name terms directly in the strings before splitting.
		uc_company_name = uc_company_name.replace("AKTIEBOLAGET", "AB")
		uc_company_name = uc_company_name.replace("HANDELSBOLAGET", "HB")
		uc_company_name = uc_company_name.replace("KOMMANDITBOLAGET", "KB")
		uc_company_name = uc_company_name.replace("AKTIEBOLAG", "AB")
		uc_company_name = uc_company_name.replace("HANDELSBOLAG", "HB")
		uc_company_name = uc_company_name.replace("KOMMANDITBOLAG", "KB")
		pdf_company_name = pdf_company_name.replace("AKTIEBOLAGET", "AB")
		pdf_company_name = pdf_company_name.replace("HANDELSBOLAGET", "HB")
		pdf_company_name = pdf_company_name.replace("KOMMANDITBOLAGET", "KB")
		pdf_company_name = pdf_company_name.replace("AKTIEBOLAG", "AB")
		pdf_company_name = pdf_company_name.replace("HANDELSBOLAG", "HB")
		pdf_company_name = pdf_company_name.replace("KOMMANDITBOLAG", "KB")

		# If there are extra spaces in the company name from the PDF but the characters are identical to UC, then use the UC version.
		if uc_company_name != pdf_company_name:
			if pdf_company_name.replace(' ', '').upper() == uc_company_name.replace(' ', '').upper():
				if pdf_company_name.isupper:
					pdf_company_name = uc_company_name.upper()
				else:
					pdf_company_name = uc_company_name

		return uc_company_name.upper().split() == pdf_company_name.upper().split()

	def personal_number_verifier(self):

		return self.UC_Data_Dictionary["uc_personnummer"] == self.PDF_Data_Dictionary["pdf_personal_number_standardized"][0]

	# UCDataExtractor().UC_Data_Dictionary["uc_personnummer"]

	def date_of_birth_verifier(self):

		uc_dob = self.UC_Data_Dictionary["uc_personnummer"][:6] # Personnummer from UC date of birth without the dash and the four final digits.
		pdf_dob = self.PDF_Data_Dictionary["pdf_personal_number_standardized"][0][:6] # Personnummer from the PDF as date of birth without the dash and the four final digits.
		pdf_sig_dob = self.PDF_Data_Dictionary["pdf_signature_dob"].replace("/", "")[2:] # Date of birth without slashes.
		# Only return true if the date of birth is the same in all three locations.
		return uc_dob == pdf_dob == pdf_sig_dob

	def pn_validity_verifier(self, pn=None):
		"""
		A function for checking the validity of Swedish personal ID Numbers, also with the final number as a control number.
		How to calculate the final control number is explained in https://www.samlogic.com/blogg/2012/11/validering-av-personnummer/
		"""
		if pn is None:
			pn = self.PDF_Data_Dictionary["pdf_personal_number_standardized"][0] # Storing the personnummer as pn.

		# If the length of the personal ID number is not between 10 and 14 characters, return False.
		if len(pn) not in range(10, 14):
			return False, False

		# If the personal ID number does not only contain digits and exactly one dash, return False.
		if not re.fullmatch(r'^\d{6}-\d{4}$', pn):
			return False, False # Invalid if it doesn't match the required format (6 digits, one dash, 4 digits)

		# If the personal ID number is not standardized, return False.
		if len(pn) != 11 and pn[-5] != '-':
			return False, False

		# If the number of digits before the dash is not exactly 6 and after the dash is not exactly 4, return False.
		numbers_before_dash = pn.split('-')[0]
		numbers_after_dash = pn.split('-')[1]
		if len(numbers_before_dash) != 6:
			return False, False
		if len(numbers_after_dash) != 4:
			return False, False

		# Continuing with the Luhn algorithm. This algorithm captures 90% of incorrect personal guarantees by extracting a control number.
		# If the personal ID number does not match the final digit of the standardized personal ID number, return False.
		# Removing the dash for the purpose of calculation.
		pn_no_dash = pn.replace("-", "")
		# Defining the sequence of multipliers for the Luhn algorithm.
		sequence = [2 if i % 2 == 0 else 1 for i in range(8)] + [2]
		# Multiplying the digits by the sequence.
		result_products = [int(pn_no_dash[i]) * sequence[i] for i in range(len(sequence))]
		# Splitting the digits from numbers greater than 9 and summing them.
		sum_digits = sum(sum(divmod(num, 10)) for num in result_products)
		# Calculating the control digits by rounding the sum to the nearest multiple of 10.
		control_digit = (math.ceil(sum_digits / 10) * 10) - sum_digits

		# Returning a boolean on whether the control figure is correct.
		return control_digit == int(pn_no_dash[-1]), control_digit

	def turned_18(self):
		from datetime import datetime as dt

		" Controlling whether the borgensman is at least 18 years old which is only performed with a long personnummer. "

		original_pn = self.PDF_Data_Dictionary['pdf_original_personnummer']
		if not original_pn:
			return None

		if PDFDataExtractor().personal_number_standardizer(original_pn)[1] == True and len(original_pn) in (12,13) and original_pn[:2] == '20':
			dob = dt.strptime(original_pn[:8], '%Y%m%d').date()
			return dt.today().date() - relativedelta(years=18) > dob
		else:
			return None

	def main_name(self):

		# Excracting the data from the class UCDataExtractor.
		personal_name = self.UC_Data_Dictionary['uc_borgensman_personal_name']

		# Use regex to extract the text within single quotes
		match = re.search(r"'(.*?)'", personal_name)
		# Check if a match is found
		if match:
			return match.group(1)
		else:
			return None
		
	# Verifier().main_name()
	
	def personal_name_verifier(self):

		# First standardize the names by turning them into uppercase.
		uc_name_uppercase = self.UC_Data_Dictionary["uc_borgensman_personal_name"].upper().replace("'", "")
		pdf_main_name_uppercase = self.PDF_Data_Dictionary['pdf_personal_name'].upper()
		pdf_signature_name_uppercase = self.PDF_Data_Dictionary['pdf_signature_name'].upper()

		# Then re-arrange the order when there is two or more names.
		if ',' in uc_name_uppercase:
			uc_name_uppercase_split = uc_name_uppercase.split(',')
			uc_name_uppercase = uc_name_uppercase_split[1].split() + ' ' + uc_name_uppercase_split[0].split()
		if len(pdf_main_name_uppercase.split(',')) >= 2:
			if ',' in pdf_main_name_uppercase:
				pdf_main_name_uppercase_split = pdf_main_name_uppercase.split(',')
				pdf_main_name_uppercase = pdf_main_name_uppercase_split[1].strip() + ' ' + pdf_main_name_uppercase_split[0].strip()
		if ',' in pdf_signature_name_uppercase:
			pdf_signature_name_uppercase_split = pdf_signature_name_uppercase.split(',')
			pdf_signature_name_uppercase = pdf_signature_name_uppercase_split[1].strip() + ' ' + pdf_signature_name_uppercase_split[0].strip()

		" Storing the main given name as capitalized. If exists. "
		if self.main_name() is not None:
			main_given_name = self.main_name()
			main_given_name_uppercase = main_given_name.upper() if main_given_name else None

		" Running the conditions. "

		# If there is an exact match in the string between UC corresponding with the PDF both in the main name and signature name, return True.
		if uc_name_uppercase == pdf_main_name_uppercase == pdf_signature_name_uppercase:
			return True

		# When the name field in the PDF is empty, return False.
		if self.PDF_Data_Dictionary['pdf_personal_name'] == "":
			return False

		# When the name field in the PDF only has 1 word, return False.
		if len(self.PDF_Data_Dictionary['pdf_personal_name'].split()) == 1:
			return False

		uc_names_uppercase_list = uc_name_uppercase.split()
		pdf_main_names_uppercase_list = pdf_main_name_uppercase.split()
		pdf_name_signature_uppercase_list = pdf_signature_name_uppercase.split()

		# When the first firstname in the 1st page of the PDF is in both UC and the PDF signature and all three have indentical lastnames.
		if (pdf_main_names_uppercase_list[0] in uc_names_uppercase_list[:-1]) and (pdf_main_names_uppercase_list[0] in pdf_name_signature_uppercase_list[:-1]) and (uc_names_uppercase_list[-1] == pdf_main_names_uppercase_list[-1] == pdf_name_signature_uppercase_list[-1]):
			return True

		# When there are middle names and only the first given name and surname are compared.
		if uc_names_uppercase_list[0] == pdf_main_names_uppercase_list[0] == pdf_name_signature_uppercase_list[0] and uc_names_uppercase_list[-1] == pdf_main_names_uppercase_list[-1] == pdf_name_signature_uppercase_list[-1]:
			return True

		# When there is a main given name present in the first names of both UC and the signature, and the surname is identical across all names.
		if self.main_name() is not None:
			if (main_given_name_uppercase in uc_names_uppercase_list and main_given_name_uppercase in pdf_name_signature_uppercase_list) and (uc_names_uppercase_list[-1] == pdf_main_names_uppercase_list[-1] == pdf_name_signature_uppercase_list[-1]):
				return True

		# An extra check for identical names.
		if uc_names_uppercase_list == pdf_main_names_uppercase_list == pdf_name_signature_uppercase_list:
			return True
		# Else none of these conditions are not upheld, return False.
		else:
			return False

	def run_verifiers(self):

		# Run all verifiers and store results in respective lists
		true_false_results = [
			self.org_nr_verifier(),
			self.company_name_verifier(),
			self.personal_number_verifier(),
			self.date_of_birth_verifier(),
			self.pn_validity_verifier()[0],
			self.personal_name_verifier()
		]

		# Storing older than 18 if the 
		if self.PDF_Data_Dictionary['pdf_original_personnummer'][0:2] == '20' and self.turned_18() == False:
			true_false_results = true_false_results + [False]

		# If all verifiers return True, return True, else return False
		return all(true_false_results), true_false_results

""" Part 5. Creating a path senser which senses files in the downloads folder and extracts the orgnr from the PDF. """

class PDFMonitor:
	
	def __init__(self, app_instance): # Store the reference to PersonalGuaranteeApp
		self.app = app_instance # Store the reference to PersonalGuaranteeApp
		self.thread = None # Store the thread reference
		# Start monitoring in a separate thread
		self.extractor_running = threading.Event()
		self.monitoring_thread = threading.Thread(target=self.monitor_folder, daemon=True)
		self.monitoring_thread.start()
		self.processing_file = False # Track if a file is being processed

	def monitor_folder(self):
		import time
		detected_path = None
		if detected_path:
			self.handle_new_file(detected_path)
		try:
			class Handler(FileSystemEventHandler):
				def on_created(_, event):
					nonlocal detected_path
					if not event.is_directory:
						detected_path = event.src_path
	
			observer = Observer()
			observer.schedule(Handler(), downloads_path, recursive=False)
			observer.start()
	
			while not detected_path:
				if not self.extractor_running.is_set():
					observer.stop()
					observer.join()
					return None
				time.sleep(1)
	
			observer.stop()
			observer.join()
			return detected_path
		except Exception as e:
			print(f"Error in monitor_folder: {e}")
			return None

	def show_popup(self, name, orgnr):
		popup_root = tk.Tk()
		popup_root.withdraw() # Hide root window
	
		popup = tk.Toplevel(popup_root)
		popup.overrideredirect(True)
		popup.attributes('-topmost', True)
		
		# Use the length of the orgnr for the width of the popup if it is longer, else use the length of the name.
		# Otherwise it creates a too small popup with short names.		
		if len(orgnr) > len(name):
			name_or_orgr = orgnr
		else:
			name_or_orgr = name
		
		width = 50 + len(name_or_orgr) * 10
		height = 50
		x = popup_root.winfo_screenwidth() // 2
		y = popup_root.winfo_screenheight() // 2
		
		popup.geometry(f"{width}x{height}+{x}+{y}")
		tk.Label(popup, text=f"{name}\n{orgnr}", font=("Arial", 16), bg="white", relief="solid").pack(fill="both", expand=True)
		
		popup.after(8000, lambda: (popup.destroy(), popup_root.destroy())) # Close both popup and root
		popup_root.mainloop()
	
	def show_cursor_notification_pdfmonitor(self, text):
		"""Shows a small notification near the cursor in system screen coordinates."""
		# Get the cursor's position directly via pyautogui (system coordinates)
		x, y = pyautogui.position()
		# Create a new Toplevel window for the notification
		notification = tk.Toplevel()
		notification.overrideredirect(True) # Remove window borders
		notification.wm_attributes("-topmost", 1) # Ensure it stays on top
		# Get window handle and apply WS_EX_LAYERED | WS_EX_TRANSPARENT
		notification.update_idletasks()
		hwnd = ctypes.windll.user32.GetParent(notification.winfo_id())
		extended_style = 0x00000020 | 0x00000080 # WS_EX_LAYERED | WS_EX_TRANSPARENT
		ctypes.windll.user32.SetWindowLongPtrW(hwnd, -20, extended_style)
		notification.geometry(f"+{x}+{y - 80}")
		# Create the label inside the Toplevel window
		label = tk.Label(notification, text=text, font=("Arial", 12, "italic"), fg="black", bg="white", 
				   relief="solid", padx=6, pady=4)
		label.pack()
		# Make the notification disappear after 2 seconds.
		notification.after(2000, notification.destroy)
		
	def handle_new_file(self, file_path):
		"""Handles a newly detected PDF file and updates the GUI."""
		if not self.processing_file: # Only update if no file is being processed
			self.processing_file = True
			print(f"Handling new file: {file_path}")
			self.app.root.after(0, self.app.update_gui)

			# self.app.update_gui() # Ensure GUI updates when a new file is processed
			# Do other actions for processing...
			self.processing_file = False # Reset flag after processing

	def orgnr_extractor(self):
		print('Running extractor...')
		import time
		while self.extractor_running.is_set():
			print("Waiting for a .tmp file to appear...")
			new_file = self.monitor_folder()
			
			if new_file and new_file.endswith(".tmp"):
				print(f"Detected {new_file}, waiting for it to disappear...")
				while os.path.exists(new_file) and self.extractor_running.is_set():
					time.sleep(0.1)
				print("Disappeared.")
			start_time = time.time()
			latest_file = None
			time.sleep(1)
			
			while self.extractor_running.is_set():
				files = glob.glob(os.path.join(downloads_path, '*'))
				matching_files = [
					file for file in files 
					if "Borgensförbindelse" in os.path.basename(file) 
					and (time.time() - os.path.getmtime(file)) < 15
				]
				if matching_files:
					latest_file = max(matching_files, key=os.path.getmtime)
					break 

				if time.time() - start_time > 10:
					latest_file = None
					break 

				time.sleep(1)

			if latest_file:
						print(f"Latest detected file: {latest_file}")

			try:
				print('Trying Orgnr Extraction.')
				reader = PdfReader(latest_file)
				all_text = "".join(page.extract_text() for page in reader.pages)
				all_text = all_text.replace('\x00','fi') # Replacing the special character with 'fi'.

				pdf_org_nr_regex = r"Org\. nr:\s*(.*?)(?=\s*Företagsnamn)"
				pdf_org_nr_re = re.search(pdf_org_nr_regex, all_text)
				pdf_org_nr_str = pdf_org_nr_re.group(1)

				pdf_personal_name_regex = r"Namn:\s*(.*)(?=\s*(?:(?:Adress)|(?:Tel)|$))"
				pdf_personal_name_re = re.search(pdf_personal_name_regex, all_text)
				pdf_personal_name_str = pdf_personal_name_re.group(1)

				# print('PDF Orgnr Str:')
				# pdf_org_nr_str
				if len(pdf_org_nr_str) >= 5 and pdf_org_nr_str[-5] == '-':
					pdf_org_nr_str = pdf_org_nr_str.replace('-', '')
				else:
					show_cursor_notification('Invalid Orgnr. Open PDF.')
				if pdf_org_nr_str.isdigit() and len(pdf_org_nr_str) == 10:
					print('Copying Orgnr.')
					pyperclip.copy(pdf_org_nr_str)
					self.show_cursor_notification_pdfmonitor('PDF Orgnr Copied.')
					self.show_popup(pdf_personal_name_str, pdf_org_nr_str)
					# After processing the PDF, now update the GUI
					self.app.root.after(0, self.app.update_gui)
					# self.app.update_gui() # Ensure GUI reflects the extracted data
			except Exception as e:
				# This will avoid unnecessary error messages if the app is turned off
				if "NoneType" in str(e):
					print("Monitoring stopped, no file to process.")
				else:
					print(f"Error extracting from PDF: {e}")

	def start_extraction(self):
		"""Start extraction only if it's not already running."""
		if not self.extractor_running.is_set():
			self.extractor_running.set()
			if self.thread is None or not self.thread.is_alive():
				self.thread = threading.Thread(target=self.orgnr_extractor, daemon=True)
				self.thread.start()
				print("PDF Monitoring started.")
				self.app.update_gui()

	def stop_extraction(self):
		"""Stops the clipboard monitoring process."""
		if self.extractor_running.is_set():
			self.extractor_running.clear() # Clear the event flag to stop any waiting threads
		# Stop the monitoring thread by setting the active flag to False
		self.monitoring_active = False
		# Ensure the monitoring thread finishes before proceeding
		if self.monitoring_thread and self.monitoring_thread.is_alive():
			self.monitoring_thread.join(timeout=1) # Wait for the thread to finish
		print("Exiting PDF Monitoring process...")

	def toggle_extractor(self):
		"""Toggle extraction on/off."""
		if self.extractor_running.is_set():
			self.stop_extraction()
			show_cursor_notification('Extraction off.')
		else:
			self.start_extraction()
			show_cursor_notification('Extraction on.')

""" Part 6. A class for buttons in the GUI. """

class Buttons:
	def __init__(self, root, page1, page2):
		self.root = root
		self.stored_data = []
		self.popup_window = None
		self.PDF_Extractor = PDFDataExtractor()
		self.UC_Extractor = UCDataExtractor()
		self.PDF_Data_Dictionary = self.PDF_Extractor.PDF_Data_Dictionary()
		self.UC_Data_Dictionary = self.UC_Extractor.UC_Data_Dictionary()
		self.gui_functions = GUIFunctions()
		self.page1 = page1
		self.page2 = page2
		self.PreviousUC = PreviousUC(self.root, self)

	def button_show_page(self, page_to_show):
		if page_to_show == self.page2:
			# Hide page 1 and show page 2
			self.page1.place_forget()
			self.page2.place(relwidth=1, relheight=1)
		else:
			# Hide page 2 and show page 1
			self.page2.place_forget()
			self.page1.place(relwidth=1, relheight=1) # Show page1 initially
			
	def button_close_window(self):
		"""A function for closing the window to run it again with a previous session."""
		if self.root:
			self.root.destroy()
			self.root = None

	def button_excel_saver(self):
		"""A function only for handling a pop-up while clicking on saving the Excel file."""
		PDF_Data_Dictionary = PDFDataExtractor().PDF_Data_Dictionary()
		if PDF_Data_Dictionary['pdf_org_nr'] == 'No PDF in Downloads':
			self.button_show_popup_non_interactive("Error. No PDF in Downloads.", 340, 40)
		else:
			result = self.gui_functions.excel_saver(self.PDF_Extractor.newest_file_extractor())
			# Check the result correctly
			if type(result) == tuple:
				if result and isinstance(result, tuple) and result[:2] == (False, False):
# =============================================================================
# 					self.button_excel_contents()
# =============================================================================
					self.button_show_popup_non_interactive("Borgen already saved in Excel.", 300, 40)
# =============================================================================
# 			elif result == True:
# 				self.button_excel_contents()
# =============================================================================
			elif result == None:
# =============================================================================
# 				self.button_excel_contents()
# =============================================================================
				self.button_show_popup_non_interactive("Error", "Saving failed. Check console for details.", 300, 80)
			
	def buttons_copy_excel_contents(self, parent, file_path):
		"""A function only for copying all the contents of the Excel file."""
		button_frame = tk.Frame(parent, bg="#f9f9f9")
		button_frame.pack(side="bottom", fill="x", pady=10, anchor = 'w')
		copy_button = tk.Button(button_frame,text="Copy All",command=lambda: self.gui_functions.copy_excel_contents(file_path),bg="#2196f3",fg="white")
		copy_button.pack(side="left", padx=10)
		
	def button_center_popup(self, popup_width, popup_height):
		screen_width, screen_height = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
		return f"{popup_width}x{popup_height}+{(screen_width - popup_width)//2}+{(screen_height - popup_height)//2}"

	def button_show_popup_non_interactive(self, message, popup_width, popup_height):
		popup = tk.Toplevel(self.root)
		popup.overrideredirect(True)
		popup.attributes('-topmost', True)
		# Center the popup on the screen
		screen_width = self.root.winfo_screenwidth()
		screen_height = self.root.winfo_screenheight()
		x_pos = (screen_width - popup_width)//2
		y_pos = (screen_height - 540 - popup_height)//2
		popup.geometry(f"{popup_width}x{popup_height}+{x_pos}+{y_pos}")
		popup.bind('<Control-Key-1>', lambda event: popup.destroy())
		tk.Label(popup, text=(message), font=("Arial", 16), bg="white", relief="solid").pack(fill="both", expand=True)
		popup.after(3000, popup.destroy)

	def button_pdf_saver(self):
		PDF_Data_Dictionary = PDFDataExtractor().PDF_Data_Dictionary()
		print(PDF_Data_Dictionary['pdf_org_nr'])
		if PDF_Data_Dictionary['pdf_org_nr'] == 'No PDF in Downloads':
			self.button_show_popup_non_interactive("Error. No PDF in Downloads.", 340, 40)
		else:
			newest_file = self.PDF_Extractor.newest_file_extractor()

			pdf_company_name = PDF_Data_Dictionary["pdf_company_name"]
			pdf_org_nr = PDF_Data_Dictionary["pdf_org_nr"]
			pdf_personal_number_standardized = PDF_Data_Dictionary["pdf_personal_number_standardized"][0]
			pdf_personal_name = PDF_Data_Dictionary["pdf_personal_name"]

			# Two target base paths
			target_paths = [
				pdf_saver_path_sharepoint,
				pdf_saver_path_desktop
			]

			# Track if any file was saved
			any_saved = False

			for base_path in target_paths:

				# Build raw filenames (without unique adjustment yet)
				orgnr_file_name_raw = os.path.join(base_path, f"{pdf_org_nr} {pdf_company_name}.pdf")
				personal_number_file_name_raw = os.path.join(base_path, f"{pdf_personal_number_standardized} {pdf_personal_name}.pdf")

				orgnr_exists = os.path.exists(orgnr_file_name_raw)
				personal_exists = os.path.exists(personal_number_file_name_raw)

				# Save only missing files
				if not orgnr_exists:
					orgnr_file_name = self.gui_functions.GUI_functions_get_unique_filename(orgnr_file_name_raw)
					shutil.copy(newest_file, orgnr_file_name)
					any_saved = True

				if not personal_exists:
					personal_number_file_name = self.gui_functions.GUI_functions_get_unique_filename(personal_number_file_name_raw)
					shutil.copy(newest_file, personal_number_file_name)
					any_saved = True

			if any_saved:
				show_cursor_notification('Saved.')
			else:
                # Both SharePoint AND Desktop copies already existed
				self.button_show_popup_non_interactive("Files already saved.", 300, 40)

	def button_excel_contents(self):
		"""A function for the button saving the data in Excel."""
		
		# Check if the popup already exists and is open
		if hasattr(self, "popup") and self.popup.winfo_exists():

			# Bring popup to front and focus it
			self.popup.lift()
			self.popup.focus_force()
			
			# Update the existing popup instead of creating a new one
			print("Popup is already open. Updating contents.")
			# Clear the previous content in the popup
			for widget in self.popup.winfo_children():
				widget.destroy()
			# Proceed with updating the popup with new content
			contents = self.gui_functions.read_excel_contents() or [['No Data Found'], []]
			headers, rows = contents[0], contents[1:]
			# Create the Treeview and buttons again inside the same popup
			first_col_width = 5
			second_col_width = 350
			tree = self.button_create_treeview(self.popup, headers, rows, first_col_width, second_col_width)
			# Create the delete button
			self.button_delete_button(self.popup, tree, rows)
			# Create the copy button
			self.buttons_copy_excel_contents(self.popup, borgens_sft_path)  # Pass actual file path
		else:
			# Create a new popup if none exists
			contents = self.gui_functions.read_excel_contents() or [['No Data Found'], []]
			headers, rows = contents[0], contents[1:]
			# Create a new popup
			self.popup = self.button_create_popup("Excel Contents", 550, max(300, 50 + len(contents) * 25))
			first_col_width = 5
			second_col_width = 350
			tree = self.button_create_treeview(self.popup, headers, rows, first_col_width, second_col_width)
			# Create the delete button
			self.button_delete_button(self.popup, tree, rows)
			# Create the copy button
			self.buttons_copy_excel_contents(self.popup, borgens_sft_path)  # Pass actual file path

	def button_create_treeview(self, parent, headers, rows, first_col_width=50, second_col_width=200):
		"""Creates a Treeview widget with left-aligned headers and minimal first column width."""
		tree = ttk.Treeview(parent, columns=headers, show="headings", height=len(rows))
		tree.pack(fill="both", expand=True)
		tree.column(headers[0], width=first_col_width, anchor="w")
		if len(headers) > 1:
			tree.column(headers[1], width=second_col_width, anchor="w")
		for header in headers:
			tree.heading(header, text=header, anchor="w")
		for row in rows:
			tree.insert("", "end", values=row)
		return tree

	def button_delete_button(self, parent, tree, rows):
		"""Adds a delete button that removes selected rows from the GUI and Excel file."""
		
		def delete_selected():
			"""Deletes selected rows from both the GUI and Excel."""
			selected_items = tree.selection()
			if not selected_items:
				show_cursor_notification("Error. Please select rows to delete.")
				return
	
			rows_to_delete = set() # Use a set to avoid duplicates
	
			# Create a fast lookup dictionary for row indices
			row_index_map = {tuple(row): i + 1 for i, row in enumerate(rows)}
	
			# Loop through selected items and get the corresponding rows
			for selected_item in selected_items:
				row_values = tuple(tree.item(selected_item)["values"]) # Convert to tuple for safe lookup
				row_to_delete = row_index_map.get(row_values) # Fast lookup instead of .index()
	
				if row_to_delete is not None:
					rows_to_delete.add(row_to_delete) # Store the index in a set
	
			# Sort the rows to delete in descending order (to avoid shifting rows)
			rows_to_delete = sorted(rows_to_delete, reverse=True) # Sort descending to avoid shifting
	
			# Call the excel_delete_row function to delete the rows from Excel
			result = True
			for row_to_delete in rows_to_delete:
				result = self.excel_delete_row(row_to_delete)
	
			# Remove any existing popups before showing a new one
			if hasattr(parent, "popup_confirm"):
				try:
					parent.popup_confirm.destroy()
					del parent.popup_confirm # Remove the reference after destroying
				except:
					pass # Ignore errors if the popup doesn't exist
	
			if hasattr(parent, "popup_error"):
				try:
					parent.popup_error.destroy()
					del parent.popup_error # Remove the reference after destroying
				except:
					pass # Ignore errors if the popup doesn't exist
			
			# After all rows have been deleted successfully, close and reload
			if result:
				# Remove the rows from the Treeview and rows list
				for selected_item in selected_items:
					tree.delete(selected_item) # Delete from Treeview
	
				# rows[:] = [row for row in rows if row not in [tree.item(selected)["values"] for selected in selected_items]]
				rows[:] = [row for idx, row in enumerate(rows) if idx + 1 not in rows_to_delete]
	
		# ✅ Now `delete_selected()` exists before we reference it in `command=`
		button_frame = tk.Frame(parent, bg="#f9f9f9")
		button_frame.pack(side="bottom", fill="x", pady=10, anchor = 'w')
		tk.Button(button_frame, text="Delete", command=delete_selected, bg="#2196f3", fg="white").pack(side="left", padx=10)
		tk.Button(button_frame, text="Close", command=parent.destroy, bg="#2196f3", fg="white").pack(side="right", padx=10)
		# Bind the Delete key to the delete_selected function
		parent.bind("<Delete>", lambda event: delete_selected())
		
	def excel_delete_row(self, row_to_delete):
		"""A function for deleting rows in Excel."""
		
		try:
			# Check if the Excel file exists
			if os.path.exists(borgens_sft_path):
				wb = openpyxl.load_workbook(borgens_sft_path)
				print(f"Workbook loaded: {borgens_sft_path}")
			else:
				print(f"Workbook not found at: {borgens_sft_path}")
				return False

			sheet = wb.active
			
			# Check if the sheet is empty
			if sheet.max_row == 1 and not sheet.cell(row=1, column=1).value:
				print("The sheet is empty, no row to delete.")
				return False
			
			row_to_delete = row_to_delete + 1
			sheet.delete_rows(row_to_delete) # Delete the row from the sheet

			# Save the changes
			wb.save(borgens_sft_path)
			print(f"Row {row_to_delete} deleted successfully.")
			show_cursor_notification('Successful Deletion.')
			return True

		except Exception as e:
			print(f"Error occurred while deleting row: {e}")
			return False

	def button_previous_UC(self):
		"""A function for using previous data from UC."""
		stored_data = self.PreviousUC.get_stored_data() # Get stored data from PreviousUC
		if not stored_data:
			print("No previous UC data available.")
			return
		# Show stored data in the popup
		self.button_create_treeview_popup("Previous UC Data", stored_data)
		
	def button_newest_borgen_name(self, chosen_path=None):
		"""
		This function extracts the newest borgen name from Downloads.
		"""
		if chosen_path is not None:
			return chosen_path  # If a path is given, return it directly.
		
		files = glob.glob(os.path.join(downloads_path, '*'))
		matching_files = [file for file in files if "Borgensförbindelse" in os.path.basename(file)]
		
		if matching_files:
			newest_file = max(matching_files, key=os.path.getmtime)
			if newest_file[-4:] != ".pdf":
				new_file = newest_file + ".pdf"
				os.rename(newest_file, new_file)
				newest_file = new_file
		else:
			newest_file = None
		
		newest_file = re.search(r"(Borgensförbindelse -.*)", newest_file).group(1)
		newest_file = newest_file.replace('.pdf','')
		show_cursor_notification('Borgen file name copied.')
		print('Borgen file name copied!')
		pyperclip.copy(newest_file)
		
	def button_copy_org_nr_no_dash(self):
		"""A function for copying the organizational number without a dash."""
		modified_org_nr = PDFDataExtractor().PDF_Data_Dictionary()["pdf_original_org_nr"].replace('-', '')
		self.root.clipboard_clear()
		self.root.clipboard_append(modified_org_nr)
		show_cursor_notification('Org.nr Copied')
		return modified_org_nr

	def button_create_popup(self, title, width, height):
		popup = tk.Toplevel(self.root)
		popup.title(title)
		popup.geometry(self.button_center_popup(width, height))
		popup.configure(bg="#f9f9f9")
		return popup

	def button_open_newest_pdf(self):
		newest_file = self.PDF_Extractor.newest_file_extractor()
		if newest_file == None:
			show_cursor_notification("No PDF In Downloads.")
		else:
			os.startfile(newest_file) # Opens with the default PDF viewer

""" Part 7. A class for the functions of the GUI. """

class GUIFunctions:
	
	def __init__(self):
		self.root = root
		self.PDF_Data_Dictionary = PDFDataExtractor().PDF_Data_Dictionary()
		self.UC_Data_Dictionary = UCDataExtractor().UC_Data_Dictionary()
	
	def excel_saver(self, newest_file):
		
		# Using regex to match everything starting from 'Borgensförbindelse -'. This is used for exporting the line to Excel.
		match = re.search(r"(Borgensförbindelse -.*)", newest_file)
		
		# Removing the .pdf ending from the string if present.
		if match:
			newest_file_shortened = match.group(1).replace('.pdf', '')
		# Removing duplicate file endings from the string if present.
		for i in range(1, 6):
			newest_file_shortened = newest_file_shortened.replace(f' ({i})', '')

		try:

			# Now open the workbook (or create a new one if not found).
			wb = openpyxl.load_workbook(borgens_sft_path) if os.path.exists(borgens_sft_path) else openpyxl.Workbook()
			sheet = wb.active # Setting sheet.
			first_empty_row = sheet.max_row + 1 # Specifying the first empty row.
			# Storing the short names for the three of us, fallback to current_user
			shortname_map = {
				'OLLESC': 'Olle',
				'olle_': 'Olle'
			}
			sft_name = shortname_map[current_user] if current_user in shortname_map else current_user

			PDF_Data_Dictionary = PDFDataExtractor().PDF_Data_Dictionary()
			
			new_row_data = [PDF_Data_Dictionary['pdf_org_nr'], newest_file_shortened, "", sft_name]
			duplicate_checker_new_row_data = new_row_data[0:2]
			
			# Check for duplicates.
			for row in sheet.iter_rows(values_only=True):
				if list(row[:4])[0:2] == duplicate_checker_new_row_data:
					return False, False
			# Write new_row_data to the first empty row.
			for col, value in enumerate(new_row_data, start=1):
				sheet.cell(row=first_empty_row, column=col, value=value)
			# Saving the updated workbook.
			wb.save(borgens_sft_path)
			show_cursor_notification('Saved.')
			return True
		# A KeyError occurs if 'pdf_org_nr' is missing from self.pdf_data.
		# This means the expected key is not found in the dictionary.
		except KeyError:
			pass
		# A PermissionError occurs if the Excel file is open in another program
		# or if the script lacks write permissions to modify the file.
		except PermissionError:
			pass
		# An OSError occurs if there is an issue with the file path, such as
		# invalid characters, the file being in use, or other OS-level problems.
		except OSError:
		# A general exception catch-all for any unexpected errors.
			pass
		except Exception:
			pass
		# Return False if error.
		return False
	
	def copy_excel_contents(self, file_path):
		"""Reads all contents from an Excel file, excluding the header row, and copies it to the clipboard."""
		try:
			df = pd.read_excel(file_path, header=None, dtype=str) # Read Excel file without setting headers
			df = df[1:] # Do not include the header.

			# Copy to clipboard while keeping empty columns
			clipboard_text = df.to_csv(sep="\t", index=False, header=False, na_rep="") # Exclude headers in output
			print(clipboard_text)
			# Create the "Copied" notification at the cursor position
			if clipboard_text != "" and clipboard_text != None:
				show_cursor_notification(text = 'Copied')
				pyperclip.copy(clipboard_text)
			else:
				show_cursor_notification(text = 'Excel Empty')
		except Exception as e:
			print("Error copying Excel contents:", e)
	
	def GUI_functions_get_unique_filename(self, file_path):
		base, extension = os.path.splitext(file_path)
		counter = 1
		while os.path.exists(file_path):
			file_path = f"{base} ({counter}){extension}"
			counter += 1
		return file_path

	def read_excel_contents(self):
		try:
			# If the file does not exist, return None.
			if not os.path.exists(borgens_sft_path):
				return None
			# Opening the workbook. The option data_only makes the extraction of values only from Excel and not Excel formulas.
			wb = openpyxl.load_workbook(borgens_sft_path, data_only=True)
			# Using the first sheet.
			sheet = wb.active
			# Reading the data by iterating the rows and reading values.
			data = [list(row[:2]) for row in sheet.iter_rows(values_only=True)]
			# Closing the workbook after the reading is completed.
			wb.close()
			# Returning the data as a list of lists.
			return data
		except Exception:
			return None

	def not_matched_text(self):
		
		Verifier_booleans = Verifier()
		bools = Verifier_booleans.run_verifiers()
		
		# The bools are in the order of:
		# 0. Orgnr
		# 1. Company name
		# 2. Personal ID number
		# 3. Date of birth
		# 4. Personal number validity [0]
		# 5. Personal name.

		orgnr_bool = bools[1][0]
		company_name_bool = bools[1][1]
		personal_ID_bool = bools[1][2]
		date_of_birth_bool = bools[1][3]
		personal_number_validity_bool = bools[1][4]
		personal_name_bool = bools[1][5]
		
		pn_validity_original_orgnr_bool = Verifier_booleans.pn_validity_verifier(self.PDF_Data_Dictionary["pdf_original_personnummer"])
		personal_number_validity_bool = Verifier_booleans.pn_validity_verifier()
		
		# Aspects that do not match are appended for page 2.
		text_list = []
		if not personal_name_bool and not personal_ID_bool and not date_of_birth_bool:
			text_list.append('The signature and the borgensman are probably not the same person.\n')
		if not company_name_bool:
			text_list.append('The company name does not match.\n')
		if not orgnr_bool:
			text_list.append('The organizational number does not match.\n')
		if not personal_name_bool:
			text_list.append('The borgensman name does not match.\n')
		if personal_number_validity_bool == (False, False):
			text_list.append('The borgensman personal ID number has an invalid format.\n')
		if (pn_validity_original_orgnr_bool and (self.PDF_Data_Dictionary["pdf_original_personnummer"][:2] == '20')):
			if not Verifier_booleans.turned_18():
				text_list.append('In the PDF, the borgensman is younger than 18 years.')
		else:
			if not personal_number_validity_bool and isinstance(personal_number_validity_bool, int):
				text_list.append(f'The borgensman personal ID number is invalid. The final digit should be {personal_number_validity_bool[1]}.\n')
			
			elif not personal_number_validity_bool[0] and isinstance(personal_number_validity_bool[1], bool):
				text_list.append('The borgensman personal ID number is incorrect.\n')
			
			if not personal_ID_bool:
				text_list.append('The personal number does not match.\n')
			
			if not date_of_birth_bool:
				text_list.append('The date of birth does not match.\n')
			
		return "".join(text_list)

	def owner_message_incorrect_borgen(self):
		text_list = []
		if not self.org_nr_verifier():
			text_list.append('Organisationsnumret (Org nr) i borgensförbindelsen har ej skrivits i enlighet med företagsregistret. ')
		if not self.personal_name_verifier() and not self.personal_number_verifier() and not self.date_of_birth_verifier():
			text_list.append('Borgensmannen och personen som signerar borgensförbindelsen måste vara samma person. ')
		else:
			if not self.personal_number_verifier():
				text_list.append('Personnumret i borgensförbindelsen har ej skrivits i enlighet med folkbokföringsregistret. ')
			if not self.company_name_verifier():
				text_list.append('Företagsnamnet i borgensförbindelsen har ej skrivits i enlighet med företagsregistret. ')
		return "".join(text_list) + 'Var vänlig och skicka in en ny borgensförbindelse.'

""" Part 8. Storing previous UC. """

class PreviousUC:
	
	"""Class for displaying and handling previous UC data in a popup window."""

	def __init__(self, root, personal_guarantee_app):
		self.root = root
		self.PersonalGuaranteeApp = personal_guarantee_app # Use existing instance
		self.stored_data = [] # Stores UC data for the current run
		self.popup_window = None
		self.tree = None

	def store_data(self, data):
		"""Stores new UC data if it doesn't already exist."""
		# Check if the data is already stored
		if data not in self.stored_data:
			if 'No Company Copied AB' not in str(data):
				self.stored_data.append(data)
				print('Data Stored.')
				print(data)
		else:
			print('Duplicate data found. Not stored.')
	
	def get_stored_data(self):
		"""Returns the stored data."""
		return self.stored_data
	
	def show_popup(self):
		"""Displays the last stored UC data in a popup window."""
		if not self.stored_data:
			print("No previous UC data available.")
			return
		
		self.create_treeview()

	def on_select(self, event=None):
		"""Handles selection of rows and copies them to clipboard."""
		selected_items = self.tree.selection()
		if not selected_items:
			return
		
		selected_row_str = [self.tree.item(item)["values"][0] for item in selected_items]
		formatted_str = "({})".format(", ".join([f"'{item}'" for item in selected_row_str]))
		pyperclip.copy(formatted_str)
		
		print(f"Copied: {formatted_str}")

	def on_double_click(self, event):
		"""Copies the double-clicked row to clipboard and updates the GUI."""
		item = self.tree.identify('item', event.x, event.y)
		if item:
			selected_row_str1 = self.tree.item(item)["values"][0]
			print('Selected Row String:', selected_row_str1)
			# Extract content inside curly braces
			selected_row_str1 = re.findall(r"\{(.*?)\}", selected_row_str1)
			# Join using " ||| " and wrap in quotes
			selected_row_str1 = f'"{" ||| ".join(selected_row_str1)}"'
			# clipboard_data = pyperclip.paste()
			print(f"Copied row: {selected_row_str1}")
			# self.PersonalGuaranteeApp.update_gui()
			self.PersonalGuaranteeApp.update_gui(selected_row_str = selected_row_str1[1:-1]) # Now using the passed reference
			# time.sleep(2)
			# pyperclip.copy(clipboard_data)

	def create_treeview(self):
		"""Creates the Treeview window to display stored UC data, ensuring only one instance exists."""
		
		# Check if the popup exists and is valid
		if hasattr(self, "popup") and self.popup and self.popup.winfo_exists():
			print("Popup is already open. Bringing it to the front.")
			self.popup.lift()  # Bring existing window to the front
			self.popup.focus_force()  # Give it focus
			return  # Exit without creating a new one
	
		# Create a new window
		self.popup = tk.Toplevel()  # Use Toplevel instead of Tk()
		self.popup.title("Previous UC Data")
	
		# Position the popup further to the right
		screen_width = self.popup.winfo_screenwidth()  # Get the width of the screen
		window_width = 1100  # Adjust the width of the window as needed
		x_position = screen_width // 2 - 320 # Position it to the left of the center (adjust as needed)
		self.popup.geometry(f"{window_width}x300+{x_position}+550")  # Set window size and position
	
		# Create a style and configure Treeview font size
		style = ttk.Style()
		style.configure("Treeview", font=("Helvetica", 14))  # Adjust font size here
		style.configure("Treeview.Heading", font=("Helvetica", 16, "bold"))  # Adjust heading size
	
		self.tree = ttk.Treeview(self.popup, columns=("Data"), show="headings", selectmode="extended")
		self.tree.pack(fill="both", expand=True, padx=10, pady=10)
	
		self.tree.heading("Data", text="Company and Personal Data")
		self.tree.column("Data", width=1100)
	
		for item in self.stored_data:
			self.tree.insert("", "end", values=(item,))
	
		self.tree.tag_configure("highlight", background="blue", foreground="white")
		self.popup.bind("<Control-c>", self.on_select)
		self.tree.bind("<Double-1>", self.on_double_click)
	
		# Handle window close properly
		self.popup.protocol("WM_DELETE_WINDOW", self.on_popup_close)
	
	def on_popup_close(self):
		"""Handles popup window closure properly."""
		if self.popup:
			self.popup.destroy()
			self.popup = None  # Reset reference so a new one can be created later

""" Part 9. Storing previous UC. """

class OtherPDF:
	
	"""Class for displaying and handling previous UC data in a popup window."""

	def __init__(self, root, personal_guarantee_app):
		self.root = root
		self.PersonalGuaranteeApp = personal_guarantee_app  # Use existing instance
		self.popup_window = None
		self.tree = None
		self.files = []
		self.matching_files = []

	def update_files(self):
		"""Updates the files and matching_files lists."""
		print("Updating files...")  # Add this line for debugging
		self.files = glob.glob(os.path.join(downloads_path, '*'))
		self.matching_files = [file for file in self.files if "Borgensförbindelse" in os.path.basename(file)]
		print(f"Files updated: {self.files}")  # Add this line to check the files

	def show_popup(self):
		"""Displays the last stored UC data in a popup window."""
		# Update the files and matching_files lists
		self.update_files()
		
		if not self.matching_files:
			print("No previous UC data available.")
			return
		
		self.create_treeview()

	def on_select(self, event=None):
		"""Handles selection of rows and copies them to clipboard."""
		borgens = self.tree.selection()
		if not borgens:
			return
		
		selected_borgen_str = [self.tree.item(item)["values"][0] for item in borgens]
		formatted_str = "({})".format(", ".join([f"'{item}'" for item in selected_borgen_str]))
		pyperclip.copy(formatted_str)
		
		print(f"Copied: {formatted_str}")

	def on_double_click(self, event):
		"""Copies the double-clicked row to the clipboard and updates the GUI."""
		item = self.tree.identify('item', event.x, event.y)
		if item:
			selected_path_str = self.tree.item(item)["values"][1]  # Get the FullPath column
			print('Selected Row Path:', selected_path_str)
			print(f"Chosen path: {selected_path_str}")
			self.PersonalGuaranteeApp.update_gui(chosen_path=selected_path_str)

	def create_treeview(self):
		"""Creates the Treeview window to display stored UC data, ensuring only one instance exists."""
		# Check if the popup exists and is valid
		if hasattr(self, "popup") and self.popup and self.popup.winfo_exists():
			print("Popup is already open. Bringing it to the front.")
			self.popup.lift()  # Bring existing window to the front
			self.popup.focus_force()  # Give it focus
			return  # Exit without creating a new one
		
		# Create a new window
		self.popup = tk.Toplevel()  # Use Toplevel instead of Tk()
		self.popup.title("Personal Guarantees in Downloads")
		
		# Position the popup further to the right
		screen_width = self.popup.winfo_screenwidth()  # Get the width of the screen
		window_width = 500  # Adjust the width of the window as needed
		x_position = screen_width // 2 + 260  # Position it to the right of the center (adjust as needed)
		self.popup.geometry(f"{window_width}x300+{x_position}+200")  # Set window size and position

		# Create a style and configure Treeview font size
		style = ttk.Style()
		style.configure("Treeview", font=("Helvetica", 14))  # Adjust font size here
		style.configure("Treeview.Heading", font=("Helvetica", 16, "bold"))  # Adjust heading size
		
		self.tree = ttk.Treeview(self.popup, columns=("Display", "FullPath"), show="headings", selectmode="extended")
		self.tree.pack(fill="both", expand=True, padx=10, pady=10)
		
		self.tree.heading("Display", text="Borgensförbindelser")
		self.tree.column("Display", width=600)
		self.tree["displaycolumns"] = ("Display")  # Hide full path from view
		
		# Sort files by modification time (newest first)
		sorted_files = sorted(self.matching_files, key=lambda f: os.path.getmtime(f), reverse=True)
		
		for full_path in sorted_files:
			filename = full_path.split("/")[-1]  # Extract only filename
			display_name = filename  # Default to full filename
			
			# Remove everything up to and including "Borgensförbindelse -"
			match = re.search(r"Borgensförbindelse -(.+)", filename)  # Keep everything after 'Borgensförbindelse -'
			if match:
				display_name = match.group(1).replace(".pdf", "")  # Remove '.pdf'
			
			self.tree.insert("", "end", values=(display_name, full_path))
		
		self.tree.tag_configure("highlight", background="blue", foreground="white")
		self.popup.bind("<Control-c>", self.on_select)
		self.tree.bind("<Double-1>", self.on_double_click)
		
		# Handle window close properly
		self.popup.protocol("WM_DELETE_WINDOW", self.on_popup_close)

	def on_popup_close(self):
		"""Handles popup window closure properly."""
		if self.popup:
			self.popup.destroy()
			self.popup = None  # Reset reference so a new one can be created later

""" Part 10. Creating a GUI. """

" The function for the GUI of the applicaton. "

class PersonalGuaranteeApp():
	def __init__(self, root):
		self.root = root
		self.extractor_running = None
		self.root.title("Personal Guarantee Controller")
		self.root.configure(bg="white")
		# Set window geometry
		self.root.geometry("1000x600")
		# Create frames
		self.page1 = tk.Frame(self.root, bg="white")
		self.page2 = tk.Frame(self.root, bg="white")
		# Pack frames (no need for place)
		self.page1.pack(fill="both", expand=True)
		self.page2.pack(fill="both", expand=True)
		# Hide page 2 by default
		self.page2.pack_forget()
		# Create an instance of Buttons to handle page switching
		self.buttons = Buttons(root, self.page1, self.page2)
		# Initialize other components
		self.GUIFunctions = GUIFunctions()
		# Initialize PDF Monitor and Verifier
		self.Verifier_ = Verifier()
		self.turned_18 = self.Verifier_.turned_18()
		self.label_borgen_page2 = None
		self.label_borgen_page1 = None
		self.label_not_matched_page2 = None
		self.label_company_name_uc = None
		self.label_company_name_pdf = None
		self.label_personnummer_uc = None
		self.label_personnummer_pdf = None
		self.label_personal_name_uc = None
		self.label_personal_name_pdf = None
		self.label_personal_name_signature = None
		self.label_dob_uc = None
		self.label_dob_pdf = None
		self.label_dob_signature = None
		self.label_organizational_number_uc = None
		self.label_organizational_number_pdf = None
		self.color_company_name = None
		self.color_personal_number = None
		self.color_personal_name = None
		self.color_date_of_birth = None
		self.color_org_nr = None
		self.label_no_company_copied_page1 = None
		self.label_no_company_copied_page2 = None
		self.label_headline_cn = None
		self.label_headline_pn = None
		self.label_headline_borgensman = None
		self.label_headline_dob = None
		self.label_headline_orgnr = None
		self.stored_data = None
		# UC Clipboard Senser
		self.clipboard_monitor_app_instance = ClipboardMonitor(self)
		# Pass the instance of PersonalGuaranteeApp (self) to PDFMonitor
		self.pdf_monitor_app_instance = PDFMonitor(self)
		self.pdf_monitor_app_instance.start_extraction()
		# Add other GUI elements
		self.create_buttons_and_labels()
		# Previous UC
		self.PreviousUC = PreviousUC(self.root, self)
		# Other PDF
		self.OtherPDF = OtherPDF(self.root, self)
		# Update GUI.
		self.update_gui()
		
	def start_pdf_monitor(self):
		"""Start PDF Monitoring from GUI."""
		print("Starting PDF Monitoring from PersonalGuaranteeApp.")
		if self.pdf_monitor_app_instance:
			self.pdf_monitor_app_instance.start_extraction()
			print('PDF Monitoring on.')

	def close_pdf_monitor(self):
		"""Stop PDF Monitoring from GUI."""
		print("Closing PDF Monitoring from PersonalGuaranteeApp.")
		if self.pdf_monitor_app_instance:
			self.pdf_monitor_app_instance.stop_extraction()
			print('PDF Monitoring closed.')
		
	def close_window(self):
		"""Close the main window (PersonalGuaranteeApp) if needed."""
		if self.root:
			self.root.destroy()
			self.root = None
			print("Main window closed!")

	def update_font_size(self, event):
		"""Dynamically adjust font size based on window width."""
		new_size = max(10, int(event.width / 40)) # Adjust divisor as needed
		self.font = ("Arial", new_size)
		self.label_company_name_uc.config(font=self.font)

	def setup_window(self):
		screen_width, screen_height = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
		x_coordinate, y_coordinate = (screen_width - self.window_width) // 2, (screen_height - self.window_height) // 2
		self.root.geometry(f"{self.window_width}x{self.window_height}+{x_coordinate}+{y_coordinate}")

	def create_frames(self):
		# Create the main frame without scrolling
		self.frame = tk.Frame(self.page1, bg="white")
		self.frame.pack(side="left", fill="both", expand=True)
		
		# Add content directly to the frame (no scrollable frame anymore)
		self.add_content()

	# No vertical line will be created or shown.
	def create_buttons_and_labels(self):
		self.buttons_create_page1()
		self.buttons_create_page2()
		self.create_labels_page1() # Ensure labels are created
		self.create_labels_page2()

	def buttons_create_page1(self):

		" Page 1. Static Buttons."
		# Export both to PDF and Excel.
		self.button_export_borgen_page1 = tk.Button(self.page1, text="Export Borgen.", command=lambda: [self.buttons.button_pdf_saver(), self.buttons.button_excel_saver()], bg="white", fg="black")
		self.button_export_borgen_page1.place(relx=0.85, rely=0.05, anchor="w")
		# Borgen saving button.
		self.button_borgen_saver_page1 = tk.Button(self.page1, text="Save Borgen PDF", command = lambda: self.buttons.button_pdf_saver(), bg="white", fg="black")
		self.button_borgen_saver_page1.place(relx=0.85, rely=0.1, anchor="w")
		# Excel saving button. Otherwise, the borgens needs to be saved manually.
		self.button_excel_saver_page1 = tk.Button(self.page1, text="Export to Excel", command = lambda: self.buttons.button_excel_saver(), bg="white", fg="black")
		self.button_excel_saver_page1.place(relx=0.85, rely=0.15, anchor="w")
		# Excel file reader button in Page 1.
		self.button_excel_reader_page1 = tk.Button(self.page1, text="Show Excel Data", command = lambda: self.buttons.button_excel_contents(), bg="white", fg="black")
		self.button_excel_reader_page1.place(relx=0.85, rely=0.20, anchor="w")
		# Refresh data in Page 1.
		self.button_refresh_page1 = tk.Button(self.page1, text="Refresh Data", command = lambda: self.update_gui(), bg="white", fg="black")
		self.button_refresh_page1.place(relx=0.85, rely=0.25, anchor="w")
		# Orgnr copy button in Page 1.
		self.button_copy_org_nr_page1 = tk.Button(self.page1, text="Copy Orgnr (PDF)", command = lambda: self.buttons.button_copy_org_nr_no_dash(), bg="white", fg="black")
		self.button_copy_org_nr_page1.place(relx=0.85, rely=0.30, anchor="w")
		# Previous UC button in Page 1.
		self.previous_UC_button_page_1 = tk.Button(self.page1, text="Previous UC", command = lambda: self.PreviousUC.create_treeview(), bg="white", fg="black")
		self.previous_UC_button_page_1.place(relx=0.85, rely=0.35, anchor="w")
		# Other PDF button in Page 1.
		self.other_PDF_page1= tk.Button(self.page1, text="Other PDF", command = lambda: self.OtherPDF.show_popup(), bg="white", fg="black")
		self.other_PDF_page1.place(relx=0.85, rely=0.4, anchor="w")
		# Open newest PDF button in Page 1.
		self.button_open_newest_pdf_page1 = tk.Button(self.page1, text="Open PDF File", command = lambda: self.buttons.button_open_newest_pdf(), bg="white", fg="black")
		self.button_open_newest_pdf_page1.place(relx=0.85, rely=0.45, anchor="w")
		# Borgen file name from downloads copy in Page 1.
		self.button_copy_borgen_name_page1 = tk.Button(self.page1, text="Borgen Name (Downloads)", command = lambda: self.buttons.button_newest_borgen_name(), bg="white", fg="black")
		self.button_copy_borgen_name_page1.place(relx=0.85, rely=0.50, anchor="w")
		
		# Page 1 to Page 2.
		self.button_to_page1_page2 = tk.Button(self.page1, text="Go to Page 2", command = lambda: self.buttons.button_show_page(self.page2), bg="white", fg="black")
		self.button_to_page1_page2.place(relx=0.08, rely=0.95, anchor="w")
		# Close app button in Page 1.
		self.close_button_page1 = tk.Button(
			self.page1, 
			text="Close Window", 
			command=lambda: (
				self.buttons.button_close_window(), 
				self.clipboard_monitor_app_instance.stop_monitoring_clipboard(),
				self.pdf_monitor_app_instance.stop_extraction()
			), 
			bg="white", 
			fg="black"
		)
		self.close_button_page1.place(relx=0.88, rely=0.95, anchor="w")

	def buttons_create_page2(self):

		" Page 2. Static Buttons."
		# Export both to PDF and Excel.
		self.button_export_borgen_page2 = tk.Button(self.page1, text="Export Borgen.", command=lambda: [self.buttons.button_pdf_saver(), self.buttons.button_excel_saver()], bg="white", fg="black")
		self.button_export_borgen_page2.place(relx=0.85, rely=0.05, anchor="w")
		# Borgen saving button.
		self.button_borgen_saver_page2 = tk.Button(self.page2, text="Save Borgen PDF", command = lambda: self.buttons.button_pdf_saver(), bg="white", fg="black")
		self.button_borgen_saver_page2.place(relx=0.85, rely=0.1, anchor="w")
		# Excel saving button. Otherwise, the borgens needs to be saved manually.
		self.button_excel_saver_page2 = tk.Button(self.page2, text="Export to Excel", command = lambda: self.buttons.button_excel_saver(), bg="white", fg="black")
		self.button_excel_saver_page2.place(relx=0.85, rely=0.15, anchor="w")
		# Excel file reader button in Page 1.
		self.button_excel_reader_page2 = tk.Button(self.page2, text="Show Excel Data", command = lambda: self.buttons.button_excel_contents(), bg="white", fg="black")
		self.button_excel_reader_page2.place(relx=0.85, rely=0.20, anchor="w")
		# Refresh data in Page 1.
		self.button_refresh_page2 = tk.Button(self.page2, text="Refresh Data", command = lambda: self.update_gui(), bg="white", fg="black")
		self.button_refresh_page2.place(relx=0.85, rely=0.25, anchor="w")
		# Orgnr copy button in Page 1.
		self.button_copy_org_nr_page2 = tk.Button(self.page2, text="Copy Orgnr (PDF)", command = lambda: self.buttons.button_copy_org_nr_no_dash(), bg="white", fg="black")
		self.button_copy_org_nr_page2.place(relx=0.85, rely=0.30, anchor="w")
		# Previous UC button in Page 2.
		self.previous_UC_button_page_2 = tk.Button(self.page2, text="Previous UC", command = lambda: self.PreviousUC.create_treeview(), bg="white", fg="black")
		self.previous_UC_button_page_2.place(relx=0.85, rely=0.35, anchor="w")
		# Other PDF button in Page 2.
		self.other_PDF_page2= tk.Button(self.page2, text="Other PDF", command = lambda: self.OtherPDF.show_popup(), bg="white", fg="black")
		self.other_PDF_page2.place(relx=0.85, rely=0.4, anchor="w")
		# Open newest PDF button in Page 2.
		self.button_open_newest_pdf_page2 = tk.Button(self.page2, text="Open PDF File", command = lambda: self.buttons.button_open_newest_pdf(), bg="white", fg="black")
		self.button_open_newest_pdf_page2.place(relx=0.85, rely=0.45, anchor="w")
		# Borgen file name from downloads copy in Page 2.
		self.button_copy_borgen_name_page2 = tk.Button(self.page2, text="Borgen Name (Downloads)", command = lambda: self.buttons.button_newest_borgen_name(), bg="white", fg="black")
		self.button_copy_borgen_name_page2.place(relx=0.85, rely=0.50, anchor="w")
		
		# Page 2 to Page 1.
		self.button_to_page2_page1 = tk.Button(self.page2, text="Go to Page 1", command = lambda: self.buttons.button_show_page(self.page1), bg="white", fg="black")
		self.button_to_page2_page1.place(relx=0.08, rely=0.95, anchor="w")
		# Close app button in Page 2.
		self.close_button_page2 = tk.Button(
			self.page2, 
			text="Close Window", 
			command=lambda: (
				self.buttons.button_close_window(), 
				self.clipboard_monitor_app_instance.stop_monitoring_clipboard(),
				self.pdf_monitor_app_instance.stop_extraction()
			), 
			bg="white", 
			fg="black"
		)
		self.close_button_page2.place(relx=0.88, rely=0.95, anchor="w")

	def create_labels_page1(self):
		" Page 1. Dynamic labels which are updated by update_gui(). "
		
		# Labels company name.
		self.label_company_name_uc = tk.Label(self.page1, text="", bg="white", font=("Arial", 15)) # UC
		self.label_company_name_uc.place(relx=0.08, rely=0.24, anchor="w")
		self.label_company_name_pdf = tk.Label(self.page1, text="", bg="white", font=("Arial", 15)) # PDF main
		self.label_company_name_pdf.place(relx=0.08, rely=0.28, anchor="w")
		# Labels personal ID number.
		self.label_personnummer_uc = tk.Label(self.page1, text="", bg="white", font=("Arial", 15)) # UC
		self.label_personnummer_uc.place(relx=0.55, rely=0.24, anchor="w")
		self.label_personnummer_pdf = tk.Label(self.page1, text="", bg="white", font=("Arial", 15)) # PDF main
		self.label_personnummer_pdf.place(relx=0.55, rely=0.28, anchor="w")
		# Labels personal name.
		self.label_personal_name_uc = tk.Label(self.page1, text="", bg="white", font=("Arial", 15)) # UC
		self.label_personal_name_uc.place(relx=0.08, rely=0.39, anchor="w")
		self.label_personal_name_pdf = tk.Label(self.page1, text="", bg="white", font=("Arial", 15)) # PDF main
		self.label_personal_name_pdf.place(relx=0.08, rely=0.44, anchor="w")
		self.label_personal_name_signature = tk.Label(self.page1, text="", bg="white", font=("Arial", 15)) # PDF signature
		self.label_personal_name_signature.place(relx=0.08, rely=0.49, anchor="w")
		# Labels date of birth.
		self.label_dob_uc = tk.Label(self.page1, text="", bg="white", font=("Arial", 15)) # UC
		self.label_dob_uc.place(relx=0.55, rely=0.39, anchor="w")
		self.label_dob_pdf = tk.Label(self.page1, text="", bg="white", font=("Arial", 15)) # PDF main
		self.label_dob_pdf.place(relx=0.55, rely=0.44, anchor="w")
		self.label_dob_signature = tk.Label(self.page1, text="", bg="white", font=("Arial", 15)) # PDF signature
		self.label_dob_signature.place(relx=0.55, rely=0.49, anchor="w")
		# Labels organizational number.
		self.label_organizational_number_uc = tk.Label(self.page1, text="", bg="white", font=("Arial", 15)) # UC
		self.label_organizational_number_uc.place(relx=0.08, rely=0.60, anchor="w")
		self.label_organizational_number_pdf = tk.Label(self.page1, text="", bg="white", font=("Arial", 15)) # PDF main
		self.label_organizational_number_pdf.place(relx=0.08, rely=0.65, anchor="w")
		# Setting headlines.
		self.label_headline_cn = tk.Label(self.page1, text="Company Name", bg="white", fg="black", font=("Arial", 20))
		self.label_headline_cn.place(relx=0.08, rely=0.18, anchor="w")
		self.label_headline_pn = tk.Label(self.page1, text="Personal ID", bg="white", fg="black", font=("Arial", 20))
		self.label_headline_pn.place(relx=0.55, rely=0.18, anchor="w")
		self.label_headline_borgensman = tk.Label(self.page1, text="Borgensman Name", bg="white", fg="black", font=("Arial", 20))
		self.label_headline_borgensman.place(relx=0.08, rely=0.34, anchor="w")
		self.label_headline_dob = tk.Label(self.page1, text="Date of Birth", bg="white", fg="black", font=("Arial", 20))
		self.label_headline_dob.place(relx=0.55, rely=0.34, anchor="w")
		self.label_headline_orgnr = tk.Label(self.page1, text="Organizational Number", bg="white", fg="black", font=("Arial", 20))
		self.label_headline_orgnr.place(relx=0.08, rely=0.55, anchor="w")
		# Labels for the second page if matched or not matched.
		self.label_borgen_page1 = tk.Label(self.page1, text="", bg="white", font=("Arial", 26), justify="left")
		self.label_borgen_page1.place(relx=0.08, rely=0.1, anchor="w")
		# Label no company copied in UC.
		self.label_no_company_copied_page1 = tk.Label(self.page1, text="", bg="white", font=("Arial", 30), justify="left") # UC
		self.label_no_company_copied_page1.place(relx=0.08, rely=0.14, anchor="w")
		# Labels showing PDF signature name when no company is copied.
		self.label_signature_name_no_company_copied_page1 = tk.Label(self.page1, text="", bg="white", font=("Arial", 20), justify="left")
		self.label_signature_name_no_company_copied_page1.place(relx=0.08, rely=0.29, anchor="w")
		# Labels showing PDF signature birthday when no company is copied.
		self.label_signature_birthday_no_company_copied_page1 = tk.Label(self.page1, text="", bg="white", font=("Arial", 20), justify="left")
		self.label_signature_birthday_no_company_copied_page1.place(relx=0.08, rely=0.35, anchor="w")
		# Labels showing PDF 1st page name when no company is copied.
		self.label_first_page_name_no_company_copied_page1 = tk.Label(self.page1, text="", bg="white", font=("Arial", 20), justify="left")
		self.label_first_page_name_no_company_copied_page1.place(relx=0.08, rely=0.41, anchor="w")
		# Labels showing PDF 1st page original personal number  when no company is copied.
		self.label_first_page_original_pn_no_company_copied_page1 = tk.Label(self.page1, text="", bg="white", font=("Arial", 20), justify="left")
		self.label_first_page_original_pn_no_company_copied_page1.place(relx=0.08, rely=0.47, anchor="w")
		# Labels showing PDF 1st page original orgnr when no company is copied.
		self.label_first_page_orgnr_no_company_copied_page1 = tk.Label(self.page1, text="", bg="white", font=("Arial", 20), justify="left")
		self.label_first_page_orgnr_no_company_copied_page1.place(relx=0.08, rely=0.53, anchor="w")
		# Labels showing PDF 1st page company name when no company is copied.
		self.label_first_page_cn__no_company_copied_page1 = tk.Label(self.page1, text="", bg="white", font=("Arial", 20), justify="left")
		self.label_first_page_cn__no_company_copied_page1.place(relx=0.08, rely=0.59, anchor="w")

	def create_labels_page2(self):
		
		" Page 2. Dynamic labels which is updated by update_gui()."

		# Label not matched.		
		self.label_not_matched_page2 = tk.Label(self.page2, text="", bg="white", fg="red", font=("Arial", 16), justify="left")
		self.label_not_matched_page2.place(relx=0.08, rely=0.3, anchor="w")
		# Labels for the second page if matched or not matched.
		self.label_borgen_page2 = tk.Label(self.page2, text="", bg="white", font=("Arial", 26), justify="left")
		self.label_borgen_page2.place(relx=0.08, rely=0.1, anchor="w")
		# Labels showing PDF 1st page when no company is copied.
		self.label_first_page_name_no_company_copied_page2 = tk.Label(self.page2, text="", bg="white", font=("Arial", 20), justify="left") # UC.
		self.label_first_page_name_no_company_copied_page2 .place(relx=0.22, rely=0.62, anchor="w")
		# Labels showing PDF 1st page original orgnr when no company is copied.
		self.label_first_page_original_pn_no_company_copied_page2= tk.Label(self.page2, text="", bg="white", font=("Arial", 20), justify="left") # UC.
		self.label_first_page_original_pn_no_company_copied_page2.place(relx=0.22, rely=0.68, anchor="w")
		# Label no company copied in UC.
		self.label_no_company_copied_page2 = tk.Label(self.page2, text="", bg="white", font=("Arial", 30), justify="left") # UC.
		self.label_no_company_copied_page2.place(relx=0.08, rely=0.14, anchor="w")
		# Labels showing PDF signature name when no company is copied.
		self.label_signature_name_no_company_copied_page2 = tk.Label(self.page2, text="", bg="white", font=("Arial", 20), justify="left") # UC.
		self.label_signature_name_no_company_copied_page2.place(relx=0.08, rely=0.29, anchor="w")
		# Labels showing PDF signature birthday when no company is copied.
		self.label_signature_birthday_no_company_copied_page2 = tk.Label(self.page2, text="", bg="white", font=("Arial", 20), justify="left") # UC.
		self.label_signature_birthday_no_company_copied_page2.place(relx=0.08, rely=0.35, anchor="w")
		# Labels showing PDF 1st page when no company is copied.
		self.label_first_page_name_no_company_copied_page2 = tk.Label(self.page2, text="", bg="white", font=("Arial", 20), justify="left") # UC.
		self.label_first_page_name_no_company_copied_page2 .place(relx=0.08, rely=0.41, anchor="w")
		# Labels showing PDF 1st page original orgnr when no company is copied.
		self.label_first_page_original_pn_no_company_copied_page2= tk.Label(self.page2, text="", bg="white", font=("Arial", 20), justify="left") # UC.
		self.label_first_page_original_pn_no_company_copied_page2.place(relx=0.08, rely=0.47, anchor="w")
		# Labels showing PDF 1st page original orgnr when no company is copied.
		self.label_first_page_orgnr_no_company_copied_page2 = tk.Label(self.page2, text="", bg="white", font=("Arial", 20), justify="left")
		self.label_first_page_orgnr_no_company_copied_page2.place(relx=0.08, rely=0.53, anchor="w")
		# Labels showing PDF 1st page company name when no company is copied.
		self.label_first_page_cn__no_company_copied_page2 = tk.Label(self.page2, text="", bg="white", font=("Arial", 20), justify="left")
		self.label_first_page_cn__no_company_copied_page2.place(relx=0.08, rely=0.59, anchor="w")

	def update_gui(self, chosen_path=None, selected_row_str=None):
		"""Schedules the GUI update to be run in the main thread."""
		self.root.after(0, self._perform_gui_update, chosen_path, selected_row_str)
	
	def _perform_gui_update(self, chosen_path=None, selected_row_str=None):
		print("update_gui() called!")
		
		# Ensure we pass chosen_path and selected_row_str to Verifier
		Verifier_booleans = Verifier(chosen_path, selected_row_str)
		Verifier_bool_all = Verifier_booleans.run_verifiers()[0]
		
		# Update UC and PDF data extraction
		self.UCDataExtractor = UCDataExtractor()
		UC_Data_Dictionary = self.UCDataExtractor.UC_Data_Dictionary(selected_row_str)
		if chosen_path:
			print('Chosen Path:',chosen_path)
		
		PDF_Data_Dictionary = PDFDataExtractor().PDF_Data_Dictionary(chosen_path)
		
		# Storing the aspects which did match as text.
		not_matched = self.GUIFunctions.not_matched_text()

		if not_matched == None: # And storing an empty string if it fully matches, with the purpose of avoiding a crash.
			not_matched = ""
		
		# Storing the booleans. The bools are in the following order.
		Verifier_booleans = Verifier(chosen_path, selected_row_str)
		bools = Verifier_booleans.run_verifiers()
		
		# The bools are in the order of:
		# 0. Orgnr
		# 1. Company name
		# 2. Personal ID number
		# 3. Date of birth
		# 4. Personal number validity [0]
		# 5. Personal name
		
		orgnr_bool = bools[1][0]
		company_name_bool = bools[1][1]
		personal_ID_bool = bools[1][2]
		date_of_birth_bool = bools[1][3]
		personal_number_validity_bool = bools[1][4]
		personal_name_bool = bools[1][5]

		if UC_Data_Dictionary['uc_original_cn'] != 'No Company Copied AB':

			self.PreviousUC.store_data(self.UCDataExtractor.update_data())

			" Page 1. "

			self.label_borgen_page1.config(text="Borgen Correct." if Verifier_bool_all else "Not Matched.",
											 fg="blue" if Verifier_bool_all else "red")
			""" Update UI labels """
			self.label_company_name_uc.config(text=f"{UC_Data_Dictionary['uc_original_cn']} (UC)")
			self.label_company_name_pdf.config(text=f"{PDF_Data_Dictionary['pdf_original_cn']} (PDF)")
			self.label_personnummer_uc.config(text=f"{UC_Data_Dictionary['uc_personnummer']} (UC)")
			self.label_personnummer_pdf.config(text=f"{PDF_Data_Dictionary['pdf_original_personnummer']} (PDF)")
			self.label_personal_name_uc.config(text=f"{UC_Data_Dictionary['uc_borgensman_personal_name']} (UC)")
			self.label_personal_name_pdf.config(text=f"{PDF_Data_Dictionary['pdf_personal_name']} (PDF 1st Page)")
			self.label_personal_name_signature.config(text=f"{PDF_Data_Dictionary['pdf_signature_name']} (PDF Signature)")

			uc_dob = UC_Data_Dictionary['uc_personnummer'][:6]
			pdf_dob = PDF_Data_Dictionary['pdf_original_personnummer_standardized'][0][:6]
			self.label_dob_uc.config(text=f"{uc_dob} (UC)")
			self.label_dob_pdf.config(text=f"{pdf_dob} (PDF 1st Page)" if PDF_Data_Dictionary['pdf_personal_number_standardized'][1] else "Invalid DOB (PDF)")
			self.label_dob_signature.config(text=f"{PDF_Data_Dictionary['pdf_signature_dob']} (PDF Signature)")
			self.label_organizational_number_uc.config(text=f"{UC_Data_Dictionary['uc_org_nr']} (UC)")
			self.label_organizational_number_pdf.config(text=f"{PDF_Data_Dictionary['pdf_org_nr']} (PDF)")

			""" Update UI colors """
			self.label_company_name_uc.config(fg="blue" if company_name_bool else "red")
			self.label_company_name_pdf.config(fg="blue" if company_name_bool else "red")
			self.label_personnummer_uc.config(fg="blue" if personal_ID_bool else "red")
			self.label_personnummer_pdf.config(fg="blue" if (personal_ID_bool and personal_number_validity_bool) else "red")
			self.label_personal_name_uc.config(fg="blue" if personal_name_bool else "red")
			self.label_personal_name_pdf.config(fg="blue" if personal_name_bool else "red")
			self.label_personal_name_signature.config(fg="blue" if personal_name_bool else "red")
			self.label_dob_uc.config(fg="blue" if date_of_birth_bool else "red")
			self.label_dob_pdf.config(fg="blue" if date_of_birth_bool else "red")
			self.label_dob_signature.config(fg="blue" if date_of_birth_bool else "red")
			self.label_organizational_number_uc.config(fg="blue" if orgnr_bool else "red")
			self.label_organizational_number_pdf.config(fg="blue" if orgnr_bool else "red")

			""" Update headlines """
			self.label_headline_cn.config(text="Company name", fg="black")
			self.label_headline_pn.config(text="Personnummer", fg="black")
			self.label_headline_borgensman.config(text="Borgensman name", fg="black")
			self.label_headline_dob.config(text="Date of birth", fg="black")
			self.label_headline_orgnr.config(text="Organizational number", fg="black")
			# When no company has been copied, update it to empty.
			self.label_no_company_copied_page1.config(text="", fg="black")
			self.label_no_company_copied_page2.config(text="", fg="black")

			# Page 1. Fixing a glitch with the white vertical line.

			self.label_borgen_page1.lift()
			self.label_company_name_uc.lift()
			self.label_company_name_pdf.lift()
			self.label_personnummer_uc.lift()
			self.label_personnummer_pdf.lift()
			self.label_personal_name_uc.lift()
			self.label_personal_name_pdf.lift()
			self.label_personal_name_signature.lift()
			self.label_dob_uc.lift()
			self.label_dob_pdf.lift()
			self.label_dob_signature.lift()
			self.label_organizational_number_uc.lift()
			self.label_organizational_number_pdf.lift()
			self.label_headline_cn.lift()
			self.label_headline_pn.lift()
			self.label_headline_borgensman.lift()
			self.label_headline_dob.lift()
			self.label_headline_orgnr.lift()
			self.label_no_company_copied_page1.lift()

			" Page 2. "

			self.label_borgen_page2.config(text="Borgen Correct." if Verifier_bool_all else "Not Matched.",
												fg="blue" if Verifier_bool_all else "red")

			self.label_not_matched_page2.config(text=not_matched, fg="blue" if Verifier_bool_all else "red")

			" With no company copied, remove the text. "
			self.label_signature_name_no_company_copied_page1.config(text="")
			self.label_signature_birthday_no_company_copied_page1.config(text="")
			self.label_first_page_name_no_company_copied_page1.config(text="")
			self.label_first_page_original_pn_no_company_copied_page1.config(text="")
			self.label_first_page_cn__no_company_copied_page1.config(text="")
			self.label_first_page_orgnr_no_company_copied_page1.config(text="")
			self.label_signature_name_no_company_copied_page2.config(text="")
			self.label_signature_birthday_no_company_copied_page2.config(text="")
			self.label_first_page_name_no_company_copied_page2.config(text="")
			self.label_first_page_original_pn_no_company_copied_page2.config(text="")
			self.label_signature_name_no_company_copied_page2.config(text="")
			self.label_signature_birthday_no_company_copied_page2.config(text="")
			self.label_first_page_name_no_company_copied_page2.config(text="")
			self.label_first_page_original_pn_no_company_copied_page2.config(text="")
			self.label_first_page_cn__no_company_copied_page2.config(text="")
			self.label_first_page_orgnr_no_company_copied_page2.config(text="")

			# Page 2. Fixing glitch with the white vertical line, always being on top.
			self.label_borgen_page2.lift()
			self.label_not_matched_page2.lift()

			# Update root.
			self.root.update_idletasks()

		else:
			
			""" Handlings case where no company is copied """
			print("No company copied!")
			
			self.label_no_company_copied_page1.config(text='No company has been copied (UC)', fg="black")
			self.label_no_company_copied_page2.config(text='No company has been copied (UC)', fg="black")
			# Fixing a glitch with the white vertical line, always being on top.
			self.label_no_company_copied_page1.lift()
			self.label_no_company_copied_page2.lift()

			if PDF_Data_Dictionary['pdf_signature_name'] != 'No PDF in Downloads':
					# When no company is copied from UC, show PDF information anyway for page 1.
					self.label_signature_name_no_company_copied_page1.config(text=f"{PDF_Data_Dictionary['pdf_signature_name']} (PDF Signature Name)", fg="blue")
					self.label_signature_birthday_no_company_copied_page1.config(text=f"{PDF_Data_Dictionary['pdf_signature_dob']} (PDF Signature Dob)", fg="blue")
					self.label_first_page_name_no_company_copied_page1.config(text=f"{PDF_Data_Dictionary['pdf_personal_name']} (PDF 1st Page Name)", fg="blue")
					self.label_first_page_original_pn_no_company_copied_page1.config(text=f"{PDF_Data_Dictionary['pdf_personal_number']} (PDF 1st Page pn)", fg="blue")
					self.label_first_page_cn__no_company_copied_page1.config(text=f"{PDF_Data_Dictionary['pdf_org_nr']} (PDF Orgnr)", fg="blue")
					self.label_first_page_orgnr_no_company_copied_page1.config(text=f"{PDF_Data_Dictionary['pdf_company_name']} (PDF Company Name)", fg="blue")

					# When no company is copied from UC, show PDF information anyway for page 2.
					self.label_signature_name_no_company_copied_page2.config(text=f"{PDF_Data_Dictionary['pdf_signature_name']} (PDF Signature Name)", fg="blue")
					self.label_signature_birthday_no_company_copied_page2.config(text=f"{PDF_Data_Dictionary['pdf_signature_dob']} (PDF Signature Dob)", fg="blue")
					self.label_first_page_name_no_company_copied_page2.config(text=f"{PDF_Data_Dictionary['pdf_personal_name']} (PDF 1st Page Name)", fg="blue")
					self.label_first_page_original_pn_no_company_copied_page2.config(text=f"{PDF_Data_Dictionary['pdf_personal_number']} (PDF 1st Page pn)", fg="blue")
					self.label_first_page_cn__no_company_copied_page2.config(text=f"{PDF_Data_Dictionary['pdf_org_nr']} (PDF Orgnr)", fg="blue")
					self.label_first_page_orgnr_no_company_copied_page2.config(text=f"{PDF_Data_Dictionary['pdf_company_name']} (PDF Company Name)", fg="blue")
	
					# Fixing a glitch with the white vertical line.
	
					self.label_signature_name_no_company_copied_page1.lift()
					self.label_signature_birthday_no_company_copied_page1.lift()
					self.label_first_page_name_no_company_copied_page1.lift()
					self.label_first_page_original_pn_no_company_copied_page1.lift()
					# When no company is copied from UC, show PDF information anyway for page 2.
					self.label_signature_name_no_company_copied_page2.lift()
					self.label_signature_birthday_no_company_copied_page2.lift()
					self.label_first_page_name_no_company_copied_page2.lift()
					self.label_first_page_original_pn_no_company_copied_page2.lift()
	
			elif PDF_Data_Dictionary['pdf_signature_name'] == 'No PDF in Downloads':
				# When no company is copied from UC, show PDF information anyway for page 1.
				self.label_signature_name_no_company_copied_page1.config(text="No PDF In Downloads.", fg="Black", font=("Arial", 24))
				self.label_signature_name_no_company_copied_page2.config(text="No PDF In Downloads.", fg="Black", font=("Arial", 24))
				# When there is no PDF, clear all the labels.
				self.label_signature_birthday_no_company_copied_page1.config(text="")
				self.label_first_page_name_no_company_copied_page1.config(text="")
				self.label_first_page_original_pn_no_company_copied_page1.config(text="")
				self.label_first_page_cn__no_company_copied_page1.config(text="")
				self.label_first_page_orgnr_no_company_copied_page1.config(text="")
				self.label_signature_name_no_company_copied_page2.config(text="")
				self.label_signature_birthday_no_company_copied_page2.config(text="")
				self.label_first_page_name_no_company_copied_page2.config(text="")
				self.label_first_page_original_pn_no_company_copied_page2.config(text="")
				self.label_first_page_cn__no_company_copied_page2.config(text="")
				self.label_first_page_orgnr_no_company_copied_page2.config(text="")

				# Fixing a glitch with the white vertical line.
				self.label_signature_name_no_company_copied_page1.lift()
				self.label_signature_birthday_no_company_copied_page1.lift()
				self.label_first_page_name_no_company_copied_page1.lift()
				self.label_first_page_original_pn_no_company_copied_page1.lift()
				self.label_signature_name_no_company_copied_page2.lift()
				self.label_signature_birthday_no_company_copied_page2.lift()
				self.label_first_page_name_no_company_copied_page2.lift()
				self.label_first_page_original_pn_no_company_copied_page2.lift()
	
			# When no company is copied, clear the regular labels.
			self.label_company_name_uc.config(text="")
			self.label_company_name_pdf.config(text="")
			self.label_personnummer_uc.config(text="")
			self.label_personnummer_pdf.config(text="")
			self.label_personal_name_uc.config(text="")
			self.label_personal_name_pdf.config(text="")
			self.label_personal_name_signature.config(text="")
			self.label_dob_uc.config(text="")
			self.label_dob_pdf.config(text="")
			self.label_dob_signature.config(text="")
			self.label_organizational_number_uc.config(text="")
			self.label_organizational_number_pdf.config(text="")
			# Clear headlines.
			self.label_headline_cn.config(text="")
			self.label_headline_pn.config(text="")
			self.label_headline_borgensman.config(text="")
			self.label_headline_dob.config(text="")
			self.label_headline_orgnr.config(text="")
			# Clear Page 2.
			self.label_borgen_page1.config(text="", fg = 'white')
			self.label_borgen_page2.config(text="", fg = 'white')
			self.label_not_matched_page2.config(text="", fg = 'white')
			
			# Update root.
			self.root.update_idletasks()

if __name__ == "__main__":
	root = tk.Tk()
	app = PersonalGuaranteeApp(root) # Starts everything, including ClipboardMonitor
	root.mainloop() # Keeps the main GUI running
