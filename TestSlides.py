
#==============================================================================================
# Import
#==============================================================================================

from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request


#==============================================================================================
# Global variables
#==============================================================================================

PRESENTATION_ID 	= '1QHusuwHBOaqzjmA-Z_q01liMabB5ewhQm5JQ0c-c9no'
ARIS_SHEET_ID 		= '1h1-MBgGrBqJPCHVBnfhFybbfmFdHljFP6oGoTPgR4-c'
BAU_SHEET_ID		= '19R5DH0wZWuntZ81RR1Pd6eUxqDbkCutKp5uW9lipcT0'
QRG_FOLDER_ID		= '1uUbfXeN3_b8ziMkarakr5-h_JvBnF-s9'


#==============================================================================================
# Functions
#==============================================================================================

#----------------------------------------------------------------------------------------------
# checks the 3 columns from the BAU to see if it is DC, RC or both.
def get_DCRC(BAU_DC_Returns_col, BAU_RC_Returns_CEN_col, BAU_RC_Returns_BRK_col):

	value = ""
	DC = False
	RC = False

	# get DC and RC values
	if BAU_DC_Returns_col == "TRUE":
		DC = True
	if BAU_RC_Returns_CEN_col == "TRUE" or BAU_RC_Returns_BRK_col == "TRUE":
		RC = True

	# if DC and RC
	if DC == True and RC == True:
		value = "DC&RC"
	
	# if DC
	if DC == True and RC == False:
		value = "DC"
	
	# if RC
	if DC == False and RC == True:
		value = "RC"
	
	return value

#----------------------------------------------------------------------------------------------
# get the system's full name in ARIS
def get_system(system):

	value = ""

	if system == "nan":
		value = "No system - this is an automated system process."
	else:
		value = system

	return(value)

#----------------------------------------------------------------------------------------------
# check if there is "other" information
def get_other(other):

	value = other

	if other.replace(" ", "") == "":
		value = "There is no additional information for this step."

	return(value)

#----------------------------------------------------------------------------------------------
# check if there is a note
def get_note(note):

	if len(note) > 0:
		value = note
	else:
		value = "There is no additional notes for this step."
	
	return(value)

#----------------------------------------------------------------------------------------------
# get the roles that came from ARIS
def get_roles(roles, bau_system):

	# roles are pulled from ARIS
	# system is pulled from BAU

	value = ""
	roles_list = []


	# if there are no roles, check if it's an interface (ITF)
	if roles == "nan" or roles == "":
		
		# if it's an interface
		if bau_system.strip() == "ITF":
			value = "Automated system process or Batch Job"
		
		# if no role and no interface
		else:
			value = "--------------------------- \n ERROR! (no role, but also not ITF) \n ---------------------------"
	
	# if there are roles
	else:
		for role in roles.splitlines():
			if role.strip() != "":
				roles_list.append(role)
		value = ", ".join(roles_list)

	return(value)

#----------------------------------------------------------------------------------------------
# check if there is a transaction code, else check if it's a manual or interface
def get_transaction_code(transaction_code, bau_system):

	value = ""

	# if there is no transaction code
	if transaction_code == "":
		# if it's an interface
		if bau_system.strip() == "ITF":
			value = "Interface"
		# if it's a manual process
		if bau_system.strip() == "MAN":
			value = "This is a manual process, therefore there is no transaction code."
	# if there is a transaction code
	else:
		value = transaction_code

	return(value)

#----------------------------------------------------------------------------------------------
# Remove existing headers
# NOTE: This is commented out as it's being done manually atm
def remove_existing_headers(input_text):

	return(input_text)

	# value = ""

	# if(input_text != ""):

	# 	split_text = input_text.strip().splitlines()
	# 	first_line = split_text[0].lower()

	# 	if "overview" in first_line or "process step" in first_line or "note" in first_line or "other" in first_line:
	# 		value = (chr(10)).join(split_text[1:]).strip()		

	# 	else:
	# 		value = input_text

	# return(value)

#----------------------------------------------------------------------------------------------
# get the process steps and split them if they are too long.
def get_process_steps(process_steps):

	line_a = ""
	line_b = ""
	split_steps = process_steps.splitlines()

	# if there is no process steps 
	if len(process_steps) == 0:
		line_a = "This is either an interface or an automated system process. Therefore there are no process steps."
	# if there are process steps
	else:
		# if the process steps are too long
		if len(split_steps) > 20:
			line_a = (chr(10)).join(split_steps[:20]).strip()
			line_b = (chr(10)).join(split_steps[20:]).strip()
		else:
			line_a = process_steps
	
	return([line_a, line_b])
	
#----------------------------------------------------------------------------------------------
# gets the row number in the BAU that has the same L4Name name. 
def get_BAU_row_number(BAU_L4Name_col, L4Name_value):

	counter = 0
	found = False

	# go through each row and compare names
	for row in BAU_L4Name_col:
		if row.lower().replace(" ", "").strip() == L4Name_value.lower().replace(" ", "").strip():
			found = True
			break
		else:
			counter += 1

	return([found, counter])

#----------------------------------------------------------------------------------------------
# checks based on ticks in the BAU whether the slide should be in the filtered document.
def check_if_slide_required_in_doc(counter, document_name, L4Name_value, BAU_L4Name_col,DC_Returns, RC_Returns_Centurion, RC_Returns_Brackenfell):

	value = False

	if counter[0] == True:

		if document_name == "DC_Returns" and DC_Returns[counter[1]] == "TRUE":
			value = True
		elif document_name == "RC_Returns_Centurion" and RC_Returns_Centurion[counter[1]] == "TRUE":
			value = True
		elif document_name == "RC_Returns_Brackenfell" and RC_Returns_Brackenfell[counter[1]] == "TRUE":
			value = True

	return(value)

#----------------------------------------------------------------------------------------------
# get all the files and folders in a folder (used to get the QRG files)
def get_files_in_folder(drive_response, drive_service, folderToLookIn_id):

	page_token = None
	files_list = []

	while True:

		response = drive_service.files().list(q="'" + folderToLookIn_id + "' in parents" ,
											spaces='drive',
											fields='nextPageToken, files(id, name, webViewLink)',
											pageToken=page_token).execute()

		# for each files/folder found
		for file in response.get('files', []):

			file_name 			= file.get('name')
			file_id 			= file.get('id')
			file_webViewLink	= file.get('webViewLink')
			
			files_list.append([file_name, file_id, file_webViewLink])

		page_token = response.get('nextPageToken', None)

		if page_token is None:
			break

	return([files_list])

#----------------------------------------------------------------------------------------------
# go down the folder directory to find the QRG files
def get_qrg_files(drive_response, drive_service):
	
	global qrg_list
	qrg_list = []

	# specific folder directory
	for L3 in get_files_in_folder(drive_response, drive_service, QRG_FOLDER_ID):
		for folder_L3 in L3:
			L4 = get_files_in_folder(drive_response, drive_service, folder_L3[1])
			if len(L4) > 0:
				for folder_L4 in L4[0]:
					artifact_folder = get_files_in_folder(drive_response, drive_service, folder_L4[1])
					if len(artifact_folder) > 0:
						for folder_artifact_folder in artifact_folder[0]:
							if "QRG" in folder_artifact_folder[0]:
								pdf_folder = get_files_in_folder(drive_response, drive_service, folder_artifact_folder[1])
								if len(pdf_folder) > 0:
									for folder_pdf_folder in pdf_folder[0]:
										if "PDF" in folder_pdf_folder[0]:
											pdf_files = get_files_in_folder(drive_response, drive_service, folder_pdf_folder[1])
											if len(pdf_files[0]) > 0:
												for pdf_file in pdf_files[0]:
													if pdf_file[0].endswith(".pdf"):


														file_fullname 	= pdf_file[0].split("_")
														qrg_location 	= file_fullname[0].strip()
														qrg_number 		= file_fullname[1].strip()
														qrg_name 		= file_fullname[-1].replace(".pdf", "").strip()
														qrg_link 		= pdf_file[2]
														qrg_list.append([qrg_location, qrg_number, qrg_name, qrg_link])

#----------------------------------------------------------------------------------------------
# checks if there is a qrg with the same name as L4, then grab the name and hyperlink(if there is one)
def get_qrg_hyperlink(L4Name):

	# if the QRG can't be found then return this
	value = ['Refer to process steps below. There is no separate QRG document for these steps.', None]

	for i in qrg_list:
		if i[2].lower().replace(" ", "") == L4Name.lower().replace(" ", ""):
			value = [i[2], i[3]]
			break

	return(value)




def get_sheet_values():

	# If modifying these scopes, delete the file token.pickle.
	print(" - Getting Scopes")
	SCOPES = 	[
					'https://www.googleapis.com/auth/presentations',
					'https://www.googleapis.com/auth/spreadsheets',
					'https://www.googleapis.com/auth/drive',
				]

	print(" - Getting document ID's")
	# document ID's


	print(" - Getting sheet ranges")
	# sheet ranges
	ARIS_EDITED_SHEET_RANGE 	= 'Manual Data Formatted Columns!A1:O'
	ARIS_SHEET_MANUALDATA_RANGE = 'Manual Data!A1:V'
	BAU_SHEET_RANGE 			= '*DC!A1:BL'


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
	# drive_service 	= get_service(secrets_path, 'https://www.googleapis.com/auth/drive', 'drive', 'v3')
	drive_service 	= build('drive', 'v3', credentials=creds)

	for document_name in ["DC_Returns", "RC_Returns_Centurion", "RC_Returns_Brackenfell"]:

		# Call the Slides API
		print(" - Calling document API's")
		
		global arisEdited_sheet_result
		global bau_sheet_result
		global presentation


		body = {'name': document_name}
		drive_response = drive_service.files().copy(fileId=PRESENTATION_ID, body=body, supportsTeamDrives=True).execute()
		presentation_copy_id = drive_response.get('id')
		

		presentation 			= service.presentations().get(presentationId=presentation_copy_id).execute()
		arisEdited_sheet_result = service_sheet.spreadsheets().values().get(spreadsheetId=ARIS_SHEET_ID, range=ARIS_EDITED_SHEET_RANGE).execute()
		aris_sheet_bau			= service_sheet.spreadsheets().values().get(spreadsheetId=ARIS_SHEET_ID, range=ARIS_SHEET_MANUALDATA_RANGE).execute()
		
		# get the rows from the sheet
		arisEdited_rows = arisEdited_sheet_result.get('values', [])
		# bau_rows 		= bau_sheet_result.get('values', [])
		aris_bau_rows 	= aris_sheet_bau.get('values', []) 
		
		# get the number for each column by using their name
		L3Name_col_num 				= arisEdited_rows[0].index("L3Name")
		L3Num_col_num 				= arisEdited_rows[0].index("L3Num")
		L3Desc_col_num 				= arisEdited_rows[0].index("L3Desc")
		L4Name_col_num 				= arisEdited_rows[0].index("L4Name")
		L4Desc_col_num 				= arisEdited_rows[0].index("L4Desc")
		Overview_col_num 			= arisEdited_rows[0].index("Overview")
		ProcessSteps_col_num 		= arisEdited_rows[0].index("Process Steps")
		Note_col_num 				= arisEdited_rows[0].index("Note")
		Other_col_num 				= arisEdited_rows[0].index("Other")
		TransactionCode_col_num 	= arisEdited_rows[0].index("Transaction Code")
		RC_col_num 					= arisEdited_rows[0].index("RC")
		DC_col_num 					= arisEdited_rows[0].index("DC")
		L4Roles_col_num 			= arisEdited_rows[0].index("L4Roles")
		L4System_col_num 			= arisEdited_rows[0].index("L4System")
		BAUNumbers_col_num 			= arisEdited_rows[0].index("BAU Numbers")

		# get the number for each column by using their name
		raw_L3Name_col_num 				= aris_bau_rows[0].index("L3Name")
		raw_L3Num_col_num 				= aris_bau_rows[0].index("L3Num")
		raw_L3Desc_col_num 				= aris_bau_rows[0].index("L3Desc")
		raw_L4Name_col_num 				= aris_bau_rows[0].index("L4Name")
		raw_L4Desc_col_num 				= aris_bau_rows[0].index("L4Desc")
		raw_Overview_col_num 			= aris_bau_rows[0].index("Overview")
		raw_ProcessSteps_col_num 		= aris_bau_rows[0].index("Process Steps")
		raw_Note_col_num 				= aris_bau_rows[0].index("Note")
		raw_Other_col_num 				= aris_bau_rows[0].index("Other")
		raw_TransactionCode_col_num 	= aris_bau_rows[0].index("Transaction Code")
		raw_RC_col_num 					= aris_bau_rows[0].index("RC")
		raw_DC_col_num 					= aris_bau_rows[0].index("DC")
		raw_L4Roles_col_num 			= aris_bau_rows[0].index("L4Roles")
		raw_L4System_col_num 			= aris_bau_rows[0].index("L4System")
		raw_BAUNumbers_col_num 			= aris_bau_rows[0].index("BAU Numbers")

		bau_ArtifactNum_col_num			= aris_bau_rows[0].index("BAU_ArtifactNum_value")
		bau_System_value_col_num		= aris_bau_rows[0].index("BAU_System_value")
		bau_TransactionCode_col_num		= aris_bau_rows[0].index("BAU_TransactionCode_value")


		# ARIS with BAU
		BAU_DC_Returns_value_col_num		= aris_bau_rows[0].index("BAU_DC_Returns_value")
		BAU_RC_Returns_CEN_value_col_num	= aris_bau_rows[0].index("BAU_RC_Returns_CEN_value")
		BAU_RC_Returns_BRK_value_col_num	= aris_bau_rows[0].index("BAU_RC_Returns_BRK_value")
		BAU_L4Name_value_col_num			= aris_bau_rows[0].index("BAU_L4Name_value")

		BAU_DC_Returns_col				= [row[BAU_DC_Returns_value_col_num] for row in aris_bau_rows]
		BAU_RC_Returns_CEN_col			= [row[BAU_RC_Returns_CEN_value_col_num] for row in aris_bau_rows]
		BAU_RC_Returns_BRK_col			= [row[BAU_RC_Returns_BRK_value_col_num] for row in aris_bau_rows]
		BAU_L4Name_col		 			= [row[BAU_L4Name_value_col_num] for row in aris_bau_rows]

		BAU_ArtifactNum_value  			= [row[bau_ArtifactNum_col_num] for row in aris_bau_rows]
		bau_System_value 				= [row[bau_System_value_col_num] for row in aris_bau_rows]
		bau_TransactionCode_value 		= [row[bau_TransactionCode_col_num] for row in aris_bau_rows]

		
		# populates the list of qrg's and their links
		get_qrg_files(drive_response, drive_service)

		print("QRG LIST")
		print(qrg_list)
		
		steps_layouts = presentation["layouts"][2]#['pageElements'];
		
		steps_layouts["layoutProperties"]["name"] = "STEPS_A"
		
		slides = presentation.get('slides')

		print("- Creating Slides")
		
		slide_counter 	= 0
		reqs 			= []
		reqs_text 		= []

		row_start 		= 1
		row_end 		= 500

		# go through each row of the sheet and create a slide
		for row in arisEdited_rows[row_start:row_end]:

			bau_row_counter = get_BAU_row_number(BAU_L4Name_col, row[L4Name_col_num])	

			if check_if_slide_required_in_doc(bau_row_counter, document_name, row[L4Name_col_num], BAU_L4Name_col, BAU_DC_Returns_col, BAU_RC_Returns_CEN_col, BAU_RC_Returns_BRK_col) == True:

				# print("Creating slide: " + str(slide_counter))

				# google slide creates a slide
				reqs.append([
				{
					'createSlide': 
					{
						'insertionIndex': str(slide_counter),
						'slideLayoutReference': 
						{
							"layoutId": 'gb7b81678ae_0_20'
						}
					}
				},
				])

				slide_counter += 1

		# update and get the presentation
		presentation = service.presentations().batchUpdate(body={'requests': reqs}, presentationId = presentation_copy_id).execute().get('replies')
		presentation = service.presentations().get(presentationId=presentation_copy_id).execute()
			


		print("- Slides Created")
		print("- Updating Slides")



		slide_counter = 0
		ITF_list = []


		# update the slides
		for row in arisEdited_rows[row_start:row_end]:

			bau_row_counter = get_BAU_row_number(BAU_L4Name_col, row[L4Name_col_num])	

			# check whether this information is required for this filtered document publish
			if check_if_slide_required_in_doc(bau_row_counter, document_name, row[L4Name_col_num], BAU_L4Name_col, BAU_DC_Returns_col, BAU_RC_Returns_CEN_col, BAU_RC_Returns_BRK_col) == True:
			
				slide_title 	= row[BAUNumbers_col_num] + " | " + get_DCRC(BAU_DC_Returns_col[bau_row_counter[1]], BAU_RC_Returns_CEN_col[bau_row_counter[1]], BAU_RC_Returns_BRK_col[bau_row_counter[1]]) + " | " + row[L4Name_col_num]
				slide_overview 	= remove_existing_headers(row[Overview_col_num])
				slide_process 	= remove_existing_headers(row[ProcessSteps_col_num])

				# check if the step is an interface(ITF)
				if bau_System_value[bau_row_counter[1]] == "ITF":
					ITF_list.append(chr(10) + slide_title + chr(10) + slide_overview + chr(10))
				
				# if not an interface
				else:
					if len(ITF_list)>0:
						slide_process = "Interface Steps: (Automated)" + chr(10) + "".join(ITF_list) + chr(10) + "Process Steps: (User based)" + chr(10) + slide_process

					ITF_list = []				

					if row[BAUNumbers_col_num] != "":

						counter = 0

						list_slide_text = [
							get_system(row[L4System_col_num]),
							get_roles(row[L4Roles_col_num], bau_System_value[bau_row_counter[1]]),
							get_transaction_code(row[TransactionCode_col_num], bau_System_value[bau_row_counter[1]]),
							slide_title,
							row[L3Name_col_num],
							get_process_steps(slide_process)[1],
							get_process_steps(slide_process)[0],
							slide_overview,
							get_qrg_hyperlink(row[L4Name_col_num]),
							remove_existing_headers(get_other(row[Other_col_num])),
							remove_existing_headers(get_note(row[Note_col_num])),
							]

						previous_slide_text = list_slide_text

						for x in presentation['slides'][slide_counter]['pageElements']:

							value = list_slide_text[counter]
							if(len(list_slide_text[counter]) == 2):
								value = list_slide_text[counter][0]

							# print(value)
								

							reqs_text.append([
							{
								'insertText':
								{
									"objectId": x['objectId'],#presentation['slides'][0]['pageElements'][0]['objectId'],
									"text": value,#"test1",
									"insertionIndex": 0
								}
							},
							])

						

							if(len(list_slide_text[counter]) == 2):
								# print(list_slide_text[counter])
								if list_slide_text[counter][1] != None:
									# if (list_slide_text[counter] != ""):

									reqs_text.append([
									{								
										
										"updateTextStyle": 
										{
											'objectId': x['objectId'],
											"style": 
											{	
												"link": 
												{
													"url": list_slide_text[counter][1]  # Please set the modified URL here.
												}
											},
											"fields": "*"
										}
									},
									])


							counter += 1



						slide_counter += 1


		presentation = service.presentations().batchUpdate(body={'requests': reqs_text}, presentationId = presentation_copy_id).execute().get('replies')
		print("- Slides Updated")

# documents_setup()
get_sheet_values()

