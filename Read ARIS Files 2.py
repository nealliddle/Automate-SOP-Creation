#-----------------------------------------------
# Import
#-----------------------------------------------

from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

import pandas as pd
import os, csv
import re
import enchant

from operator import itemgetter



#-----------------------------------------------
# Global variables
#-----------------------------------------------

BAU_folder = "G:/My Drive/Learning and Development/L&D Systems Training/BAU and Aris Processes and Activities" # master folder
d = enchant.Dict("en_US")
funct_excel_files = 	[]
rasci_excel_files = 	[]
manual_excel_files = 	[]

#-----------------------------------------------
# Definitions
#-----------------------------------------------

def documents_setup():

	# If modifying these scopes, delete the file token.pickle.
	print(" - Getting Scopes")
	SCOPES = 	[
					'https://www.googleapis.com/auth/presentations',
					'https://www.googleapis.com/auth/spreadsheets'
				]

	# document ID's
	print(" - Getting document ID's")
	BAU_SHEET_ID		= '19R5DH0wZWuntZ81RR1Pd6eUxqDbkCutKp5uW9lipcT0'

	# sheet ranges
	print(" - Getting sheet ranges")
	BAU_SHEET_RANGE 	= '*DC!A1:BL'

	creds = None

	print(" - Checking for token (pickle) ")
	if os.path.exists('token.pickle'):
		with open('token.pickle', 'rb') as token:
			creds = pickle.load(token)
	# If there are no (valid) credentials available, let the user log in.
	if not creds or not creds.valid:
		if creds and creds.expired and creds.refresh_token:
			creds.refresh(Request())
		else:
			flow = InstalledAppFlow.from_client_secrets_file(
				'client_secret.json', SCOPES)
			creds = flow.run_local_server(port=0)
		# Save the credentials for the next run
		with open('token.pickle', 'wb') as token:
			pickle.dump(creds, token)

	service 		= build('slides', 'v1', credentials=creds)
	service_sheet 	= build('sheets', 'v4', credentials=creds)

	# Call the Slides API
	print(" - Calling document API's")
	global bau_sheet_result

	bau_sheet_result 	= service_sheet.spreadsheets().values().get(spreadsheetId=BAU_SHEET_ID, range=BAU_SHEET_RANGE).execute()

def get_bau_values():
	
	bau_rows 	= bau_sheet_result.get('values', [])

	# get the number for each column by using their name
	L4Name_col_num 				= bau_rows[0].index("L4Name")
	ArtifactNum_col_num			= bau_rows[0].index("Artifact Number (Auto)")
	System_col_num				= bau_rows[0].index("System")
	TransactionCode_col_num		= bau_rows[0].index("Txs Code")
	DC_Returns_col_num 			= bau_rows[0].index("RETURNS_DC")
	RC_Returns_CEN_col_num		= bau_rows[0].index("RETURNS_RC (CEN)")
	RC_Returns_BRK_col_num		= bau_rows[0].index("BRK RC")

	global BAU_L4Name
	global BAU_ArtifactNum
	global BAU_System
	global BAU_TransactionCode
	global BAU_DC_Returns
	global BAU_RC_Returns_CEN
	global BAU_RC_Returns_BRK

	BAU_L4Name 			= [row[L4Name_col_num] for row in bau_rows]
	BAU_ArtifactNum 	= [row[ArtifactNum_col_num] for row in bau_rows]
	BAU_System			= [row[System_col_num] for row in bau_rows]
	BAU_TransactionCode = [row[TransactionCode_col_num] for row in bau_rows]
	BAU_DC_Returns		= [row[DC_Returns_col_num] for row in bau_rows]
	BAU_RC_Returns_CEN	= [row[RC_Returns_CEN_col_num] for row in bau_rows]
	BAU_RC_Returns_BRK	= [row[RC_Returns_BRK_col_num] for row in bau_rows]
	







#..................................................................
# Get files
def getFiles():

	xlsx_files_count = 0
	gsheet_files_count = 0

	# go through all the files in the folder
	for path, subdirs, files in os.walk(BAU_folder):
		for name in files:
			
			# find all excel files
			if name.endswith(".xlsx"):
				
				xlsx_files_count += 1
				
				# get all files that are in Functions folder
				if(path.split("\\")[-1] == "Functions"):

					# get exported files
					if name.endswith("Manual Copy Paste.xlsx") == False:
						funct_excel_files.append(os.path.join(path, name))

					# get manual copy-paste files
					else:
						manual_excel_files.append(os.path.join(path, name))


				# get all files that are in RASCI folder
				if(path.split("\\")[-1] == "RASCI"):
					rasci_excel_files.append(os.path.join(path, name))

			# get all google sheet files
			if name.endswith(".gsheet"):
				gsheet_files_count += 1

	print("Function files found: " + str(len(funct_excel_files)))
	print("Manual copy-paste files fount: " + str(len(manual_excel_files)))
	print("RASCI files found: " + str(len(rasci_excel_files)))
	print(xlsx_files_count)
	print(gsheet_files_count)



L3_Number = []
L3_function_name = []
L3_description = []

L4_function = []
L4_description = []
row_counter = 0



# ====================================
# MANUAL COPY PASTE
# ====================================

manual_list = []
not_transaction_code = "HEADERDETAIL,MARKOUT,EQUIPUNITID,FLOWTHRU,PICKLIST,WRICEF"

def readL4Manual_Excel_Sheet1(manual_excel_files):
	print("Reading Manual Copy-Paste Excel")

	for manual_path in manual_excel_files:

		split_path = manual_path.split("\\")[-1].split(" ")
		names = ["L3Name", "L3Num", "L3Desc", "L4Name", "L4Desc", "L4Roles", "L4System"]
		df = pd.read_excel(manual_path, sheet_name="Sheet1", engine='openpyxl', usecols="A:D", header=None, names=names)

		row_activities = 2
		row_name = 1
		for row in df.iloc[:,0].tolist():
			if row=="Activities" or row=="[-]":
				break
			else:
				row_activities += 1

		for row in df.iloc[:,0].tolist():
			if row=="Name":
				break
			else:
				row_name += 1
	
		# L3 Information
		L3Name = " ".join(split_path[2:]).split("-")[0]
		L3Num = split_path[1].replace("0", "").replace("-", ".")
		L3Desc = df.iloc[row_activities-3][0]

		for i in range(row_name,len(df)):#list((df.iloc[row_counter:len(df)])):
			row = list(df.iloc[i])

			RC = False
			DC = False
			Overview = ""
			ProcessSteps = ""
			Note = ""
			Other = ""
			Transaction_code = ""
			switch = "Overview"

			# print(row)
			if (row[0]!=""):
				# print(i)	
					
				try:
					L4Name = row[0]

					if "RC" in L4Name:
						RC = True
					if "DC" in L4Name:
						DC = True
				except:
					L4Name = ""
				
				try:
					L4Desc = row[1]
					print("-----------------------------------------")
					
					# process_list = ["process steps:", "process steps", "Process steps:", "Process steps"]
					# note_list = ["Note:", "Note", "note"]
					# other_list = 


					for line in L4Desc.splitlines():#.split(char(10)):
						line_clean = line.strip().replace(":", "").lower()


						for wrong_char in "#.,/_+()'':-*+&[ ];":
							line_clean = line_clean.replace(wrong_char, "").replace('"', "").replace("CTRL+", "").strip().lower()
						# print(line_clean)
						
						if line_clean == "overview":
							switch == "Overview"

						if line_clean[:12] == "processsteps":
							print(line_clean)
							# print("Process Step")
							switch = "Process Steps"
						
						if line_clean == "note":
							switch = "Note"
						
						if line_clean == "other":
							switch = "Other"
						
						# get transaction codes
						for word in line.split(" "):
							clean_word = word
							for wrong_char in " #.,/_+()'':-*+&[];":
								clean_word = clean_word.replace(wrong_char, "").replace('"', "").replace("CTRL+", "").strip()

							# remove words that are not transaction codes
							for bad_word in not_transaction_code.split(","):
								if bad_word in clean_word or clean_word in bad_word:

									clean_word = ""

									# print(bad_word)

							if("WRICEF58" == clean_word):
								print(clean_word)

							if clean_word == clean_word.upper() and clean_word != "" and clean_word.isdigit() == False and len(clean_word) >= 5 and d.check(clean_word) == False and clean_word not in Transaction_code and clean_word not in not_transaction_code:
								Transaction_code += clean_word + ", "

						if switch == "Overview":
							Overview += line + chr(10)
						if switch == "Process Steps":
							ProcessSteps += line + chr(10)
						if switch == "Note":
							Note += line + chr(10)
						if switch == "Other":
							Other += line + chr(10)

					if "RC" in L4Desc:
						RC = True
					if "DC" in L4Desc:
						DC = True



							# process_line.append(line)


					# for PS in process_list:
					# 	if PS in L4Desc:
					# 		ProcessSteps = re.split(PS,L4Desc)[-1].lstrip("\n").lstrip(" ").lstrip("\n").lstrip(" ")
					# 		Overview = re.split(PS,L4Desc)[0]
					# 		break

					# for note in note_list:
					# 	if note in L4Desc:
					# 		Note = re.split(grp,L4Desc)[-1].lstrip("\n").lstrip(" ").lstrip("\n").lstrip(" ")
					# 		Overview = re.split(grp,L4Desc)[0]
					# 		break

					# if ProcessSteps == "":
					# 	Overview = L4Desc

					


					# # seperate and get Process Steps and Overview
					# if "Process Steps:" in L4Desc:
					# 	ProcessSteps = re.split(r'Process Steps:',L4Desc)[-1].lstrip("\n").lstrip(" ").lstrip("\n").lstrip(" ")
					# 	Overview = re.split(r'Process Steps:|Process Steps',L4Desc)[0]
					# elif "Process Steps" in L4Desc:
					# 	ProcessSteps = re.split(r'Process Steps',L4Desc)[-1].lstrip("\n").lstrip(" ").lstrip("\n").lstrip(" ")
					# 	Overview = re.split(r'Process Steps:|Process Steps',L4Desc)[0]
					# elif "Process steps:" in L4Desc:
					# 	ProcessSteps = re.split(r'Process steps:',L4Desc)[-1].lstrip("\n").lstrip(" ").lstrip("\n").lstrip(" ")
					# 	Overview = re.split(r'Process Steps:|Process Steps|Process steps:',L4Desc)[0]
					# elif "Process steps" in L4Desc:
					# 	ProcessSteps = re.split(r'Process steps',L4Desc)[-1].lstrip("\n").lstrip(" ").lstrip("\n").lstrip(" ")
					# 	Overview = re.split(r'Process Steps:|Process Steps|Process steps',L4Desc)[0]
					# else:
					# 	print("There is NO PS")
					# 	Overview = L4Desc

					# check for RC and DC

				except:
					L4Desc = ""
				
				try:
					L4Roles = row[2]
				except:
					L4Roles = ""
				
				try:
					L4System = row[3]
				except:
					L4System = ""


	

				bau_num = ""
				num_counter = 0
				for num in BAU_L4Name:
					if num.lower().strip().replace(' ', '') == L4Name.lower().strip().replace(' ', ''):
						bau_num = 		BAU_ArtifactNum[num_counter]
						BAU_L4Name_value 			= BAU_L4Name[num_counter]
						BAU_ArtifactNum_value		= BAU_ArtifactNum[num_counter]
						BAU_System_value			= BAU_System[num_counter]
						BAU_TransactionCode_value 	= BAU_TransactionCode[num_counter]
						BAU_DC_Returns_value		= BAU_DC_Returns[num_counter]
						BAU_RC_Returns_CEN_value	= BAU_RC_Returns_CEN[num_counter]
						BAU_RC_Returns_BRK_value	= BAU_RC_Returns_BRK[num_counter]
						break



					num_counter += 1
				if bau_num == "":
					# print(L4Name.lower().strip().replace(' ', ''))
					pass

				# leadingZero_bau_num = []
				# for single_num in bau_num:
				if "-" in bau_num:
					leadingZero_bau_num = (bau_num.split("-")[0] + "-" + bau_num.split("-")[1].zfill(2))
				elif bau_num == "":
					leadingZero_bau_num = "NA"
				else:
					leadingZero_bau_num = bau_num

				# print(leadingZero_bau_num)


				# 	print
				# try:
				# 	bau_num = BAU_ArtifactNum[BAU_L4Name.index(L4Name.strip())]
				# except:
				# 	bau_num = "~ " + str(L4Name)

				manual_list.append([
					L3Name,
					L3Num,
					L3Desc,
					L4Name,
					L4Desc,
					Overview,
					ProcessSteps,
					Note,
					Other,
					Transaction_code[:-2],
					RC,
					DC,
					L4Roles,
					L4System,
					leadingZero_bau_num,
					BAU_L4Name_value,
					BAU_ArtifactNum_value,
					BAU_System_value,
					BAU_TransactionCode_value,
					BAU_DC_Returns_value,
					BAU_RC_Returns_CEN_value,
					BAU_RC_Returns_BRK_value
					])

	global manual_list_sorted
	manual_list_sorted = sorted(manual_list, key=itemgetter(14))
	

# ====================================
# Start
# ====================================

documents_setup()
get_bau_values()
getFiles()
# readL3Function_Excel_SheetA(funct_excel_files)
readL4Manual_Excel_Sheet1(manual_excel_files)

# ====================================
# Write to CSV
# ====================================

# write to csv
with open('ARIS_manual_list.csv', 'w', newline='', encoding='utf-8') as f:
	print("Exporting to CSV")

	fieldnames =[
	"L3Name",
	"L3Num",
	"L3Desc",
	"L4Name",
	"L4Desc",
	"Overview",
	"Process Steps",
	"Note",
	"Other",
	"Transaction Code",
	"RC",
	"DC",
	"L4Roles",
	"L4System",
	"BAU Numbers",
	"BAU_L4Name_value",
	"BAU_ArtifactNum_value",
	"BAU_System_value",
	"BAU_TransactionCode_value",
	"BAU_DC_Returns_value",
	"BAU_RC_Returns_CEN_value",
	"BAU_RC_Returns_BRK_value"
	]
	
	thewriter = csv.DictWriter(f, fieldnames=fieldnames)
	
	counter = 0

	# write headers
	thewriter.writerow({
		"L3Name":"L3Name",
		"L3Num":"L3Num",
		"L3Desc":"L3Desc",
		"L4Name":"L4Name",
		"L4Desc":"L4Desc",
		"Overview":"Overview",
		"Process Steps":"Process Steps",
		"Note":"Note", "Other":"Other",
		"Transaction Code":"Transaction Code",
		"RC":"RC",
		"DC":"DC",
		"L4Roles":"L4Roles",
		"L4System":"L4System",
		"BAU Numbers":"BAU Numbers",
		"BAU_L4Name_value":"BAU_L4Name_value",
		"BAU_ArtifactNum_value":"BAU_ArtifactNum_value",
		"BAU_System_value":"BAU_System_value",
		"BAU_TransactionCode_value":"BAU_TransactionCode_value",
		"BAU_DC_Returns_value":"BAU_DC_Returns_value",
		"BAU_RC_Returns_CEN_value":"BAU_RC_Returns_CEN_value",
		"BAU_RC_Returns_BRK_value":"BAU_RC_Returns_BRK_value"
		})

	for i in manual_list_sorted:
		# for x in i:
		thewriter.writerow({
			"L3Name":i[0],
			"L3Num":i[1],
			"L3Desc":i[2],
			"L4Name":i[3],
			"L4Desc":i[4],
			"Overview":i[5],
			"Process Steps":i[6],
			"Note":i[7],
			"Other":i[8],
			"Transaction Code":i[9],
			"RC":i[10],
			"DC":i[11],
			"L4Roles":i[12],
			"L4System":i[13],
			"BAU Numbers":i[14],
			"BAU_L4Name_value":i[15],
			"BAU_ArtifactNum_value":i[16],
			"BAU_System_value":i[17],
			"BAU_TransactionCode_value":i[18],
			"BAU_DC_Returns_value":i[19],
			"BAU_RC_Returns_CEN_value":i[20],
			"BAU_RC_Returns_BRK_value":i[21]
			})


