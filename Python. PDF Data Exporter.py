
# -*- coding: utf-8 -*-
"""
Created on Thu Jun 19 16:40:27 2025

@author: OLLESC
"""

"""
The Python script needs everything in the class based borgen controller except for Company Name. 
Another thing needed is also the date of the signature ,
such as DD.MM.YYYY (without the time of the day afterwards).

And it needs to pump into an Excel file with four columns.
The consecutive order should be orgnr, date signed, personal ID number, personal name.

It needs to be able to loop through a folder and output every single one, 
which at first will create duplicates for most. 
The duplicates are instead removed at the final stage of the export to the Excel file.
"""

from PyPDF2 import PdfReader
import re
import time
import glob
import os
import datetime as dt
import pandas as pd

def personal_number_orgnr_standardizer(personal_number_orgnr):
	personal_number_orgnr = str(personal_number_orgnr)
	"""
	This function takes a Swedish personal number (personal_number_orgnr) as input and 
	ensures that it follows a standard format: YYMMDD-ABCD.
	It removes any country code if present and adds a hyphen if needed.
	"""
	# Changing a space to a dash for orgnrs with a space.
	if (len(personal_number_orgnr) == 11 and personal_number_orgnr[-5] == " " and personal_number_orgnr[:2] == "55"	and personal_number_orgnr.replace(" ", "").isdigit()):  # all other chars are digits
		personal_number_orgnr = personal_number_orgnr[:6] + '-' + personal_number_orgnr[7:]
	if personal_number_orgnr:
		# If the length is not within this 10 to 13, do not standardize.
		if len(personal_number_orgnr) not in range(10,14):
			return (personal_number_orgnr, False)
		# If dash the numbers before the dash are not 6 or 8, do not standardize.
		elif ('-' in personal_number_orgnr) and len(personal_number_orgnr.split('-')[0]) not in (6,8):
			return (personal_number_orgnr, False)
		# If dash and the numbers after the dash are not 4, do not standardize.
		elif ('-' in personal_number_orgnr) and len(personal_number_orgnr.split('-')[1]) != 4:
			return (personal_number_orgnr, False)
		# If space the numbers before the space are not 6 or 8, do not standardize.
		elif (' ' in personal_number_orgnr) and len(personal_number_orgnr.split(' ')[0]) not in (6,8):
			return (personal_number_orgnr, False)
		# If space and the numbers after the space are not 4, do not standardize.
		elif (' ' in personal_number_orgnr) and len(personal_number_orgnr.split(' ')[1]) != 4:
			return (personal_number_orgnr, False)
		# If no dash and no space and the length is 11, do not standardize.
		elif ('-' not in personal_number_orgnr and ' ' not in personal_number_orgnr) and len(personal_number_orgnr) == 11:
			return (personal_number_orgnr, False)
		# If no dash and no space and the length is 13, do not standardize.
		elif ('-' not in personal_number_orgnr and ' ' not in personal_number_orgnr) and len(personal_number_orgnr) == 13:
			return (personal_number_orgnr, False)
		# If the personal_number_orgnr has a dash is in an incorrect location, do not standardize.
		elif ("-" in personal_number_orgnr and personal_number_orgnr[-5] != "-"):
			return (personal_number_orgnr, False)
		# If the personal_number_orgnr has a space is in an incorrect location, do not standardize.
		elif (" " in personal_number_orgnr and personal_number_orgnr[-5] != " "):
			return (personal_number_orgnr, False)
		# If the personal ID number is long and the person is not of legal age, do not standardize.
		elif len(personal_number_orgnr) in range(12, 14) and int(personal_number_orgnr[2:4]) > dt.datetime.now().year - 18:
			return (personal_number_orgnr, False)
		# If the personal ID number is short and the person is not of legal age, do not standardize.
		elif len(personal_number_orgnr) in range(10, 12) and int(personal_number_orgnr[0:2]) > dt.datetime.now().year - 18:
			return (personal_number_orgnr, False)
		# If the personal ID number is long, and the person is born in the 1900s and is too old, do not standardize.
		elif len(personal_number_orgnr) in range(12, 14) and personal_number_orgnr[:2].startswith("19") and int(personal_number_orgnr[:4][2:4]) < (dt.datetime.now().year - 18) % 100:  # 1900–1907 as of year 2025 → impossible:
			return (personal_number_orgnr, False)
		# If the format is short YYMMDDABCD with 10 digits and no dash, then standardize.
		elif (len(personal_number_orgnr) == 10 and personal_number_orgnr[-5] != "-"): 
			if int(personal_number_orgnr[2:4]) > 12:
				return (personal_number_orgnr, False)
			elif int(personal_number_orgnr[4:6]) > 31:
				return (personal_number_orgnr, False)
			else:
				return (personal_number_orgnr[:-4] + "-" + personal_number_orgnr[-4:], True)
		# If the format is short YYMMDD-ABCD with 10 digits and a dash, then standardize.
		elif (len(personal_number_orgnr) == 11 and personal_number_orgnr[-5] == "-"):
				if int(personal_number_orgnr[2:4]) > 12:
					return (personal_number_orgnr, False)
				elif int(personal_number_orgnr[4:6]) > 31:
					return (personal_number_orgnr, False)
				else:
					return (personal_number_orgnr, True)
		# If the format is short YYMMDD ABCD with 10 digits and a space, then standardize.
		elif (len(personal_number_orgnr) == 11 and personal_number_orgnr[-5] == " "):
				if int(personal_number_orgnr[2:4]) > 12:
					return (personal_number_orgnr, False)
				elif int(personal_number_orgnr[4:6]) > 31:
					return (personal_number_orgnr, False)
				else:
					return (personal_number_orgnr[:6] + '-' + personal_number_orgnr[7:], True)
		elif len(personal_number_orgnr) in range(12,14):
			# If the personal ID number is long and the person is not born in the 1900s or 2000s, then do not standardize the personal ID number.
			if ((len(personal_number_orgnr) == 12 or (len(personal_number_orgnr) == 13) and personal_number_orgnr[-5] == "-")) and personal_number_orgnr[:2] not in ('19', '20'):
				return (personal_number_orgnr, False)
			# If the format is long YYYYMMDDABCD with 12 digits and no dash, then standardize.
			elif (len(personal_number_orgnr) == 12 and personal_number_orgnr[-5] != "-"):
					if int(personal_number_orgnr[4:6]) > 12:
						return (personal_number_orgnr, False)
					elif int(personal_number_orgnr[6:8]) > 31:
						return (personal_number_orgnr, False)
					else:
						return (personal_number_orgnr[2:][:-4] + "-" + personal_number_orgnr[2:][-4:], True) # Excluding the first two digits signifying the century and adding a dash in the fifth spot from the last.
			# If the format is long YYYYMMDD-ABCD with 12 digits and a dash, then standardize.
			elif (len(personal_number_orgnr) == 13 and personal_number_orgnr[-5] == "-"):
					if int(personal_number_orgnr[4:6]) > 12:
						return (personal_number_orgnr, False)
					elif int(personal_number_orgnr[6:8]) > 31:
						return (personal_number_orgnr, False)
					else:
						return (personal_number_orgnr[2:], True) # Excluding the first two digits signifying the century.
			# If the format is long YYYYMMDD ABCD with 12 digits and a space, then do not standardize.
			elif (len(personal_number_orgnr) == 13 and personal_number_orgnr[-5] == " "):
					if int(personal_number_orgnr[4:6]) > 12:
						return (personal_number_orgnr, False)
					elif int(personal_number_orgnr[6:8]) > 31:
						return (personal_number_orgnr, False)
					else:
						return (personal_number_orgnr[2:8] + '-' + personal_number_orgnr[9:], True) # Excluding the first two digits signifying the century.
			else:
				return (personal_number_orgnr, False)
	else:
		return ('None', False)
	
def Data_Extractor(borgen_path):
	"""
	This function extracts text from the given PDF file.
	It returns a dictionary with extracted values.
	"""
	
	SHA_Not = False
	
	reader = PdfReader(borgen_path)
	all_text_pdf = ''
	for i, page in enumerate(reader.pages):
		if i > 1:  # Only read first 2 pages
			break
		extracted_text = page.extract_text()
		if extracted_text:
			all_text_pdf += extracted_text

	all_text_pdf = all_text_pdf.replace('\n', '')  # Remove newlines

	# --- Determine which patterns to use ---
	if 'SHA-512' in all_text_pdf and 'To review the signature validity, please open this PDF using Adobe Reader.' in all_text_pdf:
		print('SHA And')
		patterns = {
			"pdf_org_nr": r"Org\. nr:\s*([\d\-]+)(?=[A-Za-z]|: )",
			"pdf_date_signed": r"Signed:?(\d{2}\.\d{2}\.\d{4}) \d{2}:\d{2}(?=\s*eID Swedish BankID)",
			"pdf_personal_number": r"Pers\. nr:\s*([\d\-]+)(?=[A-Za-z]|: )",
			"pdf_personal_name": r"Namn:\s*([\wÅÄÖåäö ,.'-]+)"
		}
	elif 'SHA-512' not in all_text_pdf and 'To review the signature validity, please open this PDF using Adobe Reader.' in all_text_pdf:
		print('SHA Not')
		SHA_Not = True
		patterns = {
			"pdf_org_nr": r"Gäldenär(?:Pers\. nr/)?Org\. nr:\s*([^\sA-Za-z][^A-Za-z:]*)(?=[A-Za-z]|: )",
			"pdf_date_signed": r"Signed\s*(\d{2}\.\d{2}\.\d{4})",
			"pdf_personal_number": r"Borgensman(?:Pers\. nr(?:/Org\. nr)?):\s*([\d\-]+)(?=[A-Za-z]|: )",
			"pdf_personal_name": r"BorgensmanPers\. nr(?:/Org\. nr)?:\s*[\d\-]+(.*?)(?=BorgensmanPers\. nr|Gäldenär|$)"
		}
	elif 'Clicked invitation link' in all_text_pdf:
		print('Clicked invitation link')
		patterns = {
			"pdf_org_nr": r"GäldenärPers\. nr/Org\. nr:\s*([\d\-]+)(?=[A-Za-z]|: )",
			"pdf_date_signed": r"Document signed by .*?(\d{4}-\d{2}-\d{2})",
			"pdf_personal_number": r"BorgensmanPers\. nr/Org\. nr:\s*([\d\-]+)(?=[A-Za-z]|: )",
			"pdf_personal_name": r"BorgensmanPers\. nr/Org\. nr:\s*[\d\-]*\s*Namn/Firma:\s*(.*?)\s*(?:Tel\.|Adress:)"
		}
	else:
		print('Else')
		patterns = {
			"pdf_org_nr": r"Org\. nr:\s*([\d\-]+)(?=[A-Za-z]|: )",
			"pdf_date_signed": r"Signed:?(\d{2}\.\d{2}\.\d{4}) \d{2}:\d{2}(?=\s*eID Swedish BankID)",
			"pdf_personal_number": r"Pers\. nr:\s*([\d\-]+)(?=[A-Za-z]|: )",
			"pdf_personal_name": r"Namn:\s*([\wÅÄÖåäö ,.'-]+)"
		}

	# --- Extract fields ---
	# --- Extract fields ---
	data = {}

	if SHA_Not == True:
		for key, pattern in patterns.items():
			match = re.search(pattern, all_text_pdf, re.DOTALL)
			if key == "pdf_personal_name":
				if match:
					name_text = match.group(1).strip()
					# Stop at Tel, Adress, Postnr, Ort if they exist
					name_text = re.split(r'(Tel|Adress|Postnr|Ort)\s*[:.]?', name_text)[0].strip()
					data[key] = name_text
				else:
					# fallback if normal pattern fails
					fallback_match = re.search(r"(?:Namn|Namn/Firma)\s*:\s*([\wÅÄÖåäöéè ,.'-]+?)(?=Tel|Adress|Postnr|Ort|$)", all_text_pdf, re.DOTALL)
					data[key] = fallback_match.group(1).strip() if fallback_match else None
			else:
				data[key] = match.group(1).strip() if match else None
	else:
		for key, pattern in patterns.items():
			match = re.search(pattern, all_text_pdf)
			data[key] = match.group(1) if match else None
	# --- Clean the extracted values ---
	for key in ['pdf_org_nr', 'pdf_personal_number', 'pdf_personal_name']:
		if key in patterns:  # make sure the key exists
			value = data.get(key, '')  # your regex output
			if value:
				# Remove any leading headline with colon and optional spaces
				clean_value = value
				if ':' in clean_value:
					clean_value = clean_value.split(':', 1)[-1].strip()
				data[key] = clean_value
	return data

def File_Looper(excel_file_path):
	from datetime import datetime as dt
	start_time = time.time()	# ⏱ Start the stopwatch

	borgen_files = glob.glob(os.path.join(excel_file_path, '*'))
	unreadable_pdfs = []  # 🧾 Collect unreadable file paths
	table_list = []
	
	remove_prefixes = ["Namn/Firma:  ", "Tel.", "Namn:  "]

	# Add headline row
# =============================================================================
# 	table_list.append(["Orgnr", "SignedDate", "PersonalNumber", "Name", "Filename", "Pathway", "Non_Std_PN"])
# =============================================================================

	for i, file in enumerate(borgen_files, start=1):
		try:
			pdf_data_dictionary = Data_Extractor(file)
		except Exception as e:
			print(f"❌ ERROR reading PDF at index {i}: {file} - {e}")
			unreadable_pdfs.append(file)
			pdf_data_dictionary = {
				'pdf_org_nr': None,
				'pdf_date_signed': None,
				'pdf_personal_number': None,
				'pdf_personal_name': None
			}
		# Standardize and clean data
		pdf_org_nr = pdf_data_dictionary.get('pdf_org_nr')
		pdf_date_signed = pdf_data_dictionary.get('pdf_date_signed')
		pdf_personal_number = pdf_data_dictionary.get('pdf_personal_number')
		pdf_pn_non_std = pdf_personal_number
		if pdf_personal_number:
			pdf_personal_number = personal_number_orgnr_standardizer(pdf_personal_number)[0]
		pdf_personal_name = pdf_data_dictionary.get('pdf_personal_name')
		if pdf_personal_name:
			pdf_personal_name = re.sub(r'\b([A-Z])\s', r'\1', pdf_personal_name)
			for prefix in remove_prefixes:
				pdf_personal_name = pdf_personal_name.replace(prefix, "")
			
		filename = os.path.splitext(os.path.basename(file))[0]

		# Changing the format from YYYY-MM-DD into DD.MM.YYYY.
		raw_date = pdf_data_dictionary.get('pdf_date_signed')
		if raw_date:
			try:
				pdf_date_signed = dt.strptime(raw_date, "%Y-%m-%d").strftime("%d.%m.%Y")
			except ValueError:
				try:
					pdf_date_signed = dt.strptime(raw_date, "%d.%m.%Y").strftime("%d.%m.%Y")
				except ValueError:
					print(f"⚠️ Unrecognized date format: '{raw_date}' in file: {file}")
					pdf_date_signed = None
		else:
			if file not in unreadable_pdfs:
				print(f"⚠️ Missing date in file: {file}")
			pdf_date_signed = None
		
		# Always append, even if broken
		table_list.append([
			pdf_org_nr,				# [0]
			pdf_date_signed,		# [1]
			pdf_personal_number,	# [2]
			pdf_personal_name,		# [3]
			filename,				# [4]
			file,                   # [5]
			pdf_pn_non_std          # [6]
		])

		if i % 10 == 0:
			print(f"✅ Processed {i} files...")

	end_time = time.time()	# ⏹ Stop the stopwatch
	elapsed = end_time - start_time
	
	hours, rem = divmod(elapsed, 3600)
	minutes, seconds = divmod(rem, 60)
	milliseconds = (seconds % 1) * 1000
	seconds = int(seconds)
	
	print(f"\nExecution time: {int(hours)}h {int(minutes)}m {seconds}s {int(milliseconds)}ms")
	
	# Rows where any of the first four columns have None
	table_list_entries_with_none = [row for row in table_list[1:] if any(cell is None for cell in row[:4])]
	
	# Iterate over the list of lists
	for sublist in table_list_entries_with_none:
		if sublist[3] is not None:
			for prefix in remove_prefixes:
				if sublist[3].startswith(prefix):
					sublist[3] = sublist[3][len(prefix):]
					break
	
	# Rows where none of the first four columns have None
	table_list_none_removed = [row for row in table_list[1:] if all(cell is not None for cell in row[:4])]

	# Sort by date descending
	table_list_none_removed = sorted(table_list_none_removed, key=lambda x: dt.strptime(x[1], "%d.%m.%Y"), reverse=True)
	table_list_none_removed.insert(0, table_list[0])  # 👑 Add header back

	# Remove "Företagsnamn" from first value in each row
	for row in table_list:
		if row[0]:
			row[0] = row[0].replace("Företagsnamn", "").strip()

	for row in table_list_none_removed:
		if row[0]:
			row[0] = row[0].replace("Företagsnamn", "").strip()

	return table_list_none_removed, table_list, table_list_entries_with_none, unreadable_pdfs


all_results = []

cleaned_borgens_20260216 = File_Looper("C:/Users/olle_/OneDrive/Skrivbord/Portfolio/Personal Guarantees")


# Collect only variables that are defined
for var_name in [
	"cleaned_borgens_20260216",
]:
	if var_name in globals():
		all_results.extend(globals()[var_name][0])

# =============================================================================
# # Collect only variables that are defined
# for var_name in [
# 	"results_borgens_2020",
# 	"results_borgens_2021",
# 	"results_borgens_2022",
# 	"results_borgens_2023",
# 	"results_borgens_2024",
# 	"results_borgens_2025"
# ]:
# 	if var_name in globals():
# 		all_results.extend(globals()[var_name][0])
# =============================================================================

from datetime import datetime as dt

def parse_signed_date(value):
	# Skip empty or non-string values
	if not value or not isinstance(value, str):
		return None

	# Quick sanity check: real dates contain dots
	if "." not in value:
		print("Skipping non-date value:", value)
		return None

	try:
		parsed = dt.strptime(value, "%d.%m.%Y")
		return parsed
	except ValueError:
		print("Invalid date:", value)
		return None


# Example usage:
rows = [
	{"SignedDate": "31.01.2020"},
	{"SignedDate": "SignedDate"},
	{"SignedDate": ""},
	{"SignedDate": "15.06.2021"},
]

for row in rows:
	value = row["SignedDate"]
	parsed = parse_signed_date(value)
	print("Input:", value, " => Parsed:", parsed)


# -----------------------------------------
# FIXED NORMALIZATION BLOCK
# -----------------------------------------

for sublist in all_results:
	date_str = sublist[1]

	# Skip invalid values entirely
	if not isinstance(date_str, str) or date_str.strip() == "":
		print("Skipping non-date value:", date_str)
		continue

	try:
		if "-" in date_str:  # YYYY-MM-DD
			date_obj = dt.strptime(date_str, "%Y-%m-%d")
		else:  # DD.MM.YYYY
			date_obj = dt.strptime(date_str, "%d.%m.%Y")

		# Store normalized date
		sublist[1] = date_obj.strftime("%d.%m.%Y")

	except ValueError:
		print("Invalid date in all_results:", date_str)
		continue
	
	
# Sort by (date, orgnr, name)
#   → Newest date first
#   → Orgnr ascending
#   → Name ascending
# Skip the top row (header)
data_rows = all_results[1:]

def safe_date(value):
	# If the value is not a proper DD.MM.YYYY date, return a minimal date
	try:
		return dt.strptime(value, "%d.%m.%Y")
	except Exception:
		print("Skipping invalid date during sort:", value)
		# Return a very old date so invalid rows end up last
		return dt.min

data_rows.sort(
	key=lambda x: (
		safe_date(x[1]),                    # date
		x[0],                               # orgnr
		x[2].lower() if len(x) > 2 else ""  # name (safe)
	),
	reverse=True
)

# Put header back on top
all_results = [all_results[0]] + data_rows
# Capitalizing every name, moving first name behind comme and removing comma, 
for sublist in all_results:
	name = sublist[3].strip()
	
	# Remove "Adress" if present
	name = name.replace("Adress", "")
	# Removing double spacing.
	name = name.replace("  ", "")
	# If the name has a comma, swap Last, First → First Last
	if ',' in name:
		parts = [part.strip() for part in name.split(',')]
		name = ' '.join(parts[::-1])
	
	# Capitalize each word
	name = ' '.join(word.capitalize() for word in name.split())
	
	sublist[3] = name

# Replace backslashes with slashes in the final value of each sublist
for sublist in all_results:
	if len(sublist) > 0 and isinstance(sublist[-1], str):
		sublist[-1] = sublist[-1].replace("\\", "/")

# Create a DataFrame
df = pd.DataFrame(all_results, columns=['Orgnr', 'SignedDate', 'PersonalNumber', 'Name', 'Filename','Pathway','Non_Std_Pn'])

# Pathway.
export_path = os.path.join(os.path.expanduser("~"), "Desktop", "C:/Users/olle_/OneDrive/Skrivbord/Portfolio/Personal Guarantees.xlsx")

# Export to Excel
df.to_excel(export_path, index=False)

print(f"File saved to {export_path}")
