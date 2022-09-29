' 'Required for statistical purposes==========================================================================================
' name_of_script = "NOTICES - APPOINTMENT LETTER.vbs"
' start_time = timer
' STATS_counter = 1                          'sets the stats counter at one
' STATS_manualtime = 195                               'manual run time in seconds
' STATS_denomination = "C"       'C is for each CASE
' 'END OF stats block=========================================================================================================
' 
'Because we are running these locally, we are going to get rid of all the calls to GitHub...
if func_lib_run <> true then 
	FuncLib_URL = "I:\Blue Zone Scripts\Functions Library.vbs"
	Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
	Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
	text_from_the_other_script = fso_command.ReadAll
	fso_command.Close
	Execute text_from_the_other_script
	func_lib_run = true
end if
'END FUNCTIONS LIBRARY BLOCK================================================================================================
' 
' ' 'CHANGELOG BLOCK ===========================================================================================================
' ' 'Starts by defining a changelog array
' ' changelog = array()
' ' 
' ' 'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
' ' 'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
' ' call changelog_update("12/6/2016", "Corrected bug which was leaving appointment time off of case notes for in office interviews.", "Charles Potter, DHS")
' ' call changelog_update("11/28/2016", "Enabled access to Hennepin County users. Added TIKL, and added variables to allow DAIL scrubber support. Updated error message handling within dialog.", "Ilse Ferris, Hennepin County")
' ' call changelog_update("11/20/2016", "Initial version.", "Ilse Ferris, Hennepin County")
' ' 
' ' 'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
' ' changelog_display
' ' 'END CHANGELOG BLOCK =======================================================================================================
' 
' 'CLASSES----------------------------------------------------------------------------------------------------------------------
' 'IF THIS WORKS, CONSIDER INCORPORATING INTO FUNCTIONS LIBRARY
' 
' 'The following defines a class called "address" which carries several simple address properties, which can be used by scripts.
' class address
'     public street           'Defines a "street" property.
'     public city             'Defines a "city" property.
'     public state            'Defines a "state" property.
'     public zip              'Defines a "zip" property.
' 
'     'Creates a "oneline" property containing the entire address on a single line.
'     public property get oneline
'         oneline = street & ", " & city & ", " & state & " " & zip
'     end property
' 
'     'Creates a "twolines" property containing the entire address on two lines, split into an array.
'     public property get twolines
'         twolines = array(street, city & ", " & state & " " & zip)
'     end property
' end class
' 
' 
' 'Declaring variables needed by the script
' 'First, determining the county code. If it isn't declared, it will ask (proxy)
' 
' get_county_code
' 
' 
' if worker_county_code = "x101" then
'     agency_office_array = array("Aitkin")
' elseif worker_county_code = "x102" then
'     agency_office_array = array("Anoka", "Blaine", "Columbia Heights", "Lexington")
' elseif worker_county_code = "x103" then
'     agency_office_array = array("Becker")
' elseif worker_county_code = "x104" then
'     agency_office_array = array("Beltrami")
' elseif worker_county_code = "x105" then
'     agency_office_array = array("Benton")
' elseif worker_county_code = "x106" then
'     script_end_procedure("You have NOT defined an intake address with Veronica Cary. Have an alpha user email Veronica Cary and provide your in-person intake address. The script will now stop.")
' elseif worker_county_code = "x107" then
'     agency_office_array = array("Blue Earth")
' elseif worker_county_code = "x108" then
'     agency_office_array = array("New Ulm", "Sleepy Eye", "Springfield")
' elseif worker_county_code = "x109" then
'     agency_office_array = array("Cloquet", "Moose Lake")
' elseif worker_county_code = "x110" then
'     agency_office_array = array("Carver")
' elseif worker_county_code = "x111" then
'     agency_office_array = array("Cass")
' elseif worker_county_code = "x112" then
'     agency_office_array = array("Chippewa")
' elseif worker_county_code = "x113" then
'     agency_office_array = array("Center City", "North Branch")
' elseif worker_county_code = "x114" then
'     agency_office_array = array("Clay")
' elseif worker_county_code = "x115" then
'     agency_office_array = array("Clearwater")
' elseif worker_county_code = "x116" then
'     agency_office_array = array("Cook")
' elseif worker_county_code = "x117" then
'     agency_office_array = array("Cottonwood")
' elseif worker_county_code = "x118" then
'     agency_office_array = array("Crow Wing")
' elseif worker_county_code = "x119" then
'     agency_office_array = array("Dakota")
' elseif worker_county_code = "x120" then
'     agency_office_array = array("Dodge") 'MNPrairie County Alliance is Dodge, Steele & Waseca Counties
'     elseif worker_county_code = "x121" then
'     agency_office_array = array("Douglas")
' elseif worker_county_code = "x122" then
'     agency_office_array = array("Faribault")
' elseif worker_county_code = "x123" then
'     agency_office_array = array("Fillmore")
' elseif worker_county_code = "x124" then
'     agency_office_array = array("Freeborn")
' elseif worker_county_code = "x125" then
'     agency_office_array = array("Goodhue")
' elseif worker_county_code = "x126" then
'     agency_office_array = array("Grant")
' elseif worker_county_code = "x127" then
'     agency_office_array = array("South Minneapolis", "Northwest", "South Suburban", "North Hub", "West", "Central/NE")
' elseif worker_county_code = "x128" then
'     agency_office_array = array("Houston")
' elseif worker_county_code = "x129" then
'     agency_office_array = array("Hubbard")
' elseif worker_county_code = "x130" then
'     agency_office_array = array("Isanti")
' elseif worker_county_code = "x131" then
'     agency_office_array = array("Itasca")
' elseif worker_county_code = "x132" then
'     agency_office_array = array("Jackson")
' elseif worker_county_code = "x133" then
'     agency_office_array = array("Kanabec")
' elseif worker_county_code = "x134" then
'     agency_office_array = array("Kandiyohi")
' elseif worker_county_code = "x135" then
'     agency_office_array = array("Kittson")
' elseif worker_county_code = "x136" then
'     agency_office_array = array("Koochiching")
' elseif worker_county_code = "x137" then
'     agency_office_array = array("Lac Qui Parle")
' elseif worker_county_code = "x138" then
'     agency_office_array = array("Lake")
' elseif worker_county_code = "x139" then
'     agency_office_array = array("Lake of the Woods")
' elseif worker_county_code = "x140" then
'     agency_office_array = array("Le Sueur")
' elseif worker_county_code = "x141" then
'     agency_office_array = array("Lincoln")
' elseif worker_county_code = "x142" then
'     agency_office_array = array("Lyon")
' elseif worker_county_code = "x143" then
'     agency_office_array = array("Mcleod")
' elseif worker_county_code = "x144" then
'     agency_office_array = array("Mahnomen")
' elseif worker_county_code = "x145" then
'     agency_office_array = array("Marshall")
' elseif worker_county_code = "x146" then
'     agency_office_array = array("Martin")
' elseif worker_county_code = "x147" then
'     agency_office_array = array("Meeker")
' elseif worker_county_code = "x148" then
'     agency_office_array = array("Mille Lacs")
' elseif worker_county_code = "x149" then
'     agency_office_array = array("Morrison")
' elseif worker_county_code = "x150" then
'     agency_office_array = array("Mower")
' elseif worker_county_code = "x151" then
'     agency_office_array = array("Murray")
' elseif worker_county_code = "x152" then
'     agency_office_array = array("St. Peter", "North Mankato")
' elseif worker_county_code = "x153" then
'     agency_office_array = array("Nobles")
' elseif worker_county_code = "x154" then
'     agency_office_array = array("Norman")
' elseif worker_county_code = "x155" then
'     agency_office_array = array("Olmsted")
' elseif worker_county_code = "x156" then
'     agency_office_array = array("Otter Tail")
' elseif worker_county_code = "x157" then
'     agency_office_array = array("Pennington")
' elseif worker_county_code = "x158" then
'     agency_office_array = array("Pine City", "Sandstone")
' elseif worker_county_code = "x159" then
'     agency_office_array = array("Pipestone")
' elseif worker_county_code = "x160" then
'     agency_office_array = array("Crookston", "Fosston")
' elseif worker_county_code = "x161" then
'     agency_office_array = array("Pope")
' elseif worker_county_code = "x162" then
'     agency_office_array = array("Ramsey", "Fairview", "AIFC", "CAC-Bigelow", "Midway", "North St.Paul") 'adding more locations to Ramsey County
' elseif worker_county_code = "x163" then
'     agency_office_array = array("Red Lake")
' elseif worker_county_code = "x164" then
'     agency_office_array = array("Redwood")
' elseif worker_county_code = "x165" then
'     agency_office_array = array("Renville")
' elseif worker_county_code = "x166" then
'     agency_office_array = array("Rice")
' elseif worker_county_code = "x167" then
'     agency_office_array = array("Rock")
' elseif worker_county_code = "x168" then
'     agency_office_array = array("Roseau")
' elseif worker_county_code = "x169" then
'     agency_office_array = array("Duluth", "Virginia", "Hibbing", "Ely")
' elseif worker_county_code = "x170" then
'     agency_office_array = array("Scott")
' elseif worker_county_code = "x171" then
'     agency_office_array = array("Sherburne")
' elseif worker_county_code = "x172" then
'     agency_office_array = array("Sibley")
' elseif worker_county_code = "x173" then
'     agency_office_array = array("St. Cloud", "Melrose")
' elseif worker_county_code = "x174" then
'     agency_office_array = array("Owatonna", "Waseca", "Mantorville") 'MNPrairie County Alliance is Dodge, Steele & Waseca Counties
' elseif worker_county_code = "x175" then
'     agency_office_array = array("Stevens")
' elseif worker_county_code = "x176" then
'     agency_office_array = array("Swift")
' elseif worker_county_code = "x177" then
'     agency_office_array = array("Long Prairie", "Staples")
' elseif worker_county_code = "x178" then
'     agency_office_array = array("Traverse")
' elseif worker_county_code = "x179" then
'     agency_office_array = array("Wabasha")
' elseif worker_county_code = "x180" then
'     agency_office_array = array("Wadena")
' elseif worker_county_code = "x181" then
'     agency_office_array = array("Waseca") 'MNPrairie County Alliance is Dodge, Steele & Waseca Counties
'    elseif worker_county_code = "x182" then
'     agency_office_array = array("Cottage Grove", "Forest Lake", "Stillwater", "Woodbury")
' elseif worker_county_code = "x183" then
'     agency_office_array = array("Watonwan")
' elseif worker_county_code = "x184" then
'     agency_office_array = array("Wilkin")
' elseif worker_county_code = "x185" then
'     agency_office_array = array("Winona")
' elseif worker_county_code = "x186" then
'     agency_office_array = array("Wright")
' elseif worker_county_code = "x187" then
'     agency_office_array = array("Yellow Medicine")
' elseif worker_county_code = "x188" then
'     script_end_procedure("You have NOT defined an intake address with Veronica Cary. Have an alpha user email Veronica Cary and provide your in-person intake address. The script will now stop.")
' elseif worker_county_code = "x192" then
'     agency_office_array = array("Detroit Lakes", "Naytahwaush", "Bagley", "Mahnomen")
' end if
' '
' 
' county_office_list = ""     'Blanking this out because it may contain old info from the old global variables (from before this was integrated in this script)
' 
' call convert_array_to_droplist_items(agency_office_array, county_office_list)
' 
' 'DIALOGS----------------------------------------------------------------------------------------------------
' 'NOTE: this dialog contains a special modification to allow dynamic creation of the county office list. You cannot edit it in
' '   Dialog Editor without modifying the commented line.
' BeginDialog appointment_letter_dialog, 0, 0, 156, 355, "Appointment letter"
'   EditBox 75, 5, 50, 15, MAXIS_case_number
'   DropListBox 50, 25, 95, 15, "new application"+chr(9)+"recertification", app_type
'   CheckBox 10, 43, 150, 10, "Check here if this is a reschedule.", reschedule_check
'   EditBox 50, 55, 95, 15, CAF_date
'   CheckBox 10, 75, 130, 10, "Check here if this is a recert and the", no_CAF_check
'   DropListBox 70, 100, 75, 15, "Select one..."+chr(9)+"PHONE"+chr(9)+county_office_list, interview_location     'This line dynamically creates itself based on the information in the FUNCTIONS FILE.
'   EditBox 70, 120, 75, 15, interview_date
'   EditBox 70, 140, 75, 15, interview_time
'   EditBox 80, 160, 65, 15, client_phone
'   CheckBox 10, 200, 95, 10, "Client appears expedited", expedited_check
'   CheckBox 10, 215, 135, 10, "Same day interview offered/declined", same_day_declined_check
'   EditBox 10, 250, 135, 15, expedited_explanation
'   CheckBox 10, 280, 135, 10, "Check here if you left V/M with client", voicemail_check
'   EditBox 85, 305, 60, 15, worker_signature
'   ButtonGroup ButtonPressed
'     OkButton 25, 325, 50, 15
'     CancelButton 85, 325, 50, 15
'   Text 25, 10, 50, 10, "Case number:"
'   Text 15, 30, 30, 10, "App type:"
'   Text 15, 60, 35, 10, "CAF date:"
'   Text 30, 85, 105, 10, "CAF hasn't been received yet."
'   Text 15, 105, 55, 10, "Int'vw location:"
'   Text 15, 125, 50, 10, "Interview date: "
'   Text 15, 145, 50, 10, "Interview time:"
'   Text 15, 160, 60, 20, "Client phone (if phone interview):"
'   GroupBox 5, 185, 145, 85, "Expedited questions"
'   Text 10, 230, 135, 20, "If expedited interview date is not within six days of the application, explain:"
'   Text 45, 290, 75, 10, "requesting a call back."
'   Text 15, 310, 65, 10, "Worker signature:"
' EndDialog
' 
' 'Case number only dialog for x127 users
' BeginDialog case_number_dialog, 0, 0, 136, 60, "Case number dialog"
'   EditBox 60, 10, 60, 15, MAXIS_case_number
'   ButtonGroup ButtonPressed
'     OkButton 15, 35, 50, 15
'     CancelButton 70, 35, 50, 15
'   Text 10, 15, 45, 10, "Case number:"
' EndDialog
' 
' 'Hennepin County appointment letter
' BeginDialog Hennepin_appt_dialog, 0, 0, 296, 75, "Hennepin County appointment letter"
'   EditBox 205, 25, 55, 15, interview_date
'   EditBox 65, 50, 115, 15, worker_signature
'   ButtonGroup ButtonPressed
'     OkButton 185, 50, 50, 15
'     CancelButton 240, 50, 50, 15
'   EditBox 65, 25, 55, 15, CAF_date
'   Text 5, 55, 60, 10, "Worker signature:"
'   Text 140, 30, 60, 10, "Appointment date:"
'   GroupBox 20, 10, 255, 35, "Enter a new appointment date only if it's a date county offices are not open."
'   Text 30, 30, 35, 10, "CAF date:"
' EndDialog
' 
' 'THE SCRIPT----------------------------------------------------------------------------------------------------
' 'Connects to BlueZone & gathers case number
' EMConnect ""
' call MAXIS_case_number_finder(MAXIS_case_number)
' 
' If worker_county_code = "x127" then
' 	Do
' 		Do
' 			err_msg = ""
' 			dialog case_number_dialog
' 			If ButtonPressed = 0 then stopscript
' 			If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbnewline & "* Enter a valid case number."
' 			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
' 		Loop until err_msg = ""
' 		call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
' 	LOOP UNTIL are_we_passworded_out = false
' 
' 	'grabs CAF date, turns CAF date into string for variable
' 	call autofill_editbox_from_MAXIS(HH_member_array, "PROG", CAF_date)
' 	CAF_date = CAF_date & ""
' 
' 	'creates interview date for 7 calendar days from the CAF date
' 	interview_date = dateadd("d", 7, CAF_date)
' 	If interview_date < date then interview_date = dateadd("d", 7, date)
' 	interview_date = interview_date & ""		'turns interview date into string for variable
' 
' 	'Establishing values for variables that do not appear in the x127 dialog
' 	app_type = "new application"
' 	interview_location = "PHONE"
' 	interview_time = "9:00 AM - 1:00 PM"
' 
' 	Do
' 		Do
'     		err_msg = ""
'     		dialog Hennepin_appt_dialog
'     		cancel_confirmation
' 			If isdate(CAF_date) = False then err_msg = err_msg & vbnewline & "* Enter a valid CAF date."
'     		If isdate(interview_date) = False then err_msg = err_msg & vbnewline & "* Enter a valid interview date."
'     		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
'     	Loop until err_msg = ""
'     	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
'     LOOP UNTIL are_we_passworded_out = false
' Else
' 	'This Do...loop shows the appointment letter dialog, and contains logic to require most fields.
' 	Do
' 		Do
' 			err_msg = ""
' 			Dialog appointment_letter_dialog
' 			If ButtonPressed = cancel then stopscript
' 			If isnumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbnewline & "* You must fill in a valid case number. Please try again."
' 			CAF_date = replace(CAF_date, ".", "/")
' 			If no_CAF_check = checked and app_type = "new application" then no_CAF_check = unchecked 'Shuts down "no_CAF_check" so that it will validate the date entered. New applications can't happen if a CAF wasn't provided.
' 			If no_CAF_check = unchecked and isdate(CAF_date) = False then err_msg = err_msg & vbnewline & "* You did not enter a valid CAF date (MM/DD/YYYY format). Please try again."
' 	    	If interview_location = "Select one..." then err_msg = err_msg & vbnewline & "* You must select an interview location! Please try again!"
' 	    	If interview_location = "PHONE" and client_phone = "" then err_msg = err_msg & vbnewline & "* If this is a phone interview, you must enter a phone number! Please try again."
' 	    	interview_date = replace(interview_date, ".", "/")
' 	    	If isdate(interview_date) = False then err_msg = err_msg & vbnewline & "* You did not enter a valid interview date (MM/DD/YYYY format). Please try again."
' 	    	If interview_time = "" then err_msg = err_msg & vbnewline & "* You must type an interview time. Please try again."
' 	    	If no_CAF_check = checked then exit do 'If no CAF was turned in, this layer of validation is unnecessary, so the script will skip it.
' 	    	If expedited_check = checked and datediff("d", CAF_date, interview_date) > 6 and expedited_explanation = "" then err_msg = err_msg & vbnewline & "* You have indicated that your case is expedited, but scheduled the interview date outside of the six-day window. An explanation of why is required for QC purposes."
' 			If worker_signature = "" then err_msg = err_msg & vbnewline & "* You must provide a signature for your case note."
' 			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
' 		Loop until err_msg = ""
' 		call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
' 	LOOP UNTIL are_we_passworded_out = false
' END IF
' 
' 'Creates a variable to contain the agency addresses. "Address" is a class defined above.
' set agency_address = new address
' 
' 'As these are all MN intake locations, the state for all of them will be MN.
' agency_address.state = "MN"
' 
' IF interview_location = "Anoka" THEN
'     agency_address.street = "2100 3rd Ave, Suite 400"
'     agency_address.city = "Anoka"
'     agency_address.zip = "55303"
' ELSEIF interview_location = "Blaine" THEN
'     agency_address.street = "1201 89th Ave, Suite 400"
'     agency_address.city = "Blaine"
'     agency_address.zip = "55434"
' ELSEIF interview_location = "Columbia Heights" THEN
'     agency_address.street = "3980 Central Ave NE"
'     agency_address.city = "Columbia Heights"
'     agency_address.zip = "55421"
' ELSEIF interview_location = "Lexington" THEN
'     agency_address.street = "9201 S. HWY Drive, Suite B"
'     agency_address.city = "Lexington"
'     agency_address.zip = "55014"
' END IF
' 
' 'This is a temporary MsgBox that expires 09/01/2015. It is designed to "make sure" that the address is correct. Because this function is new, I want to be ABSOLUTELY SURE it's working before notices get sent out.
' If interview_location <> "PHONE" and datediff("d", date, #9/1/2015#) > 0 then
' 	double_check_MsgBox = MsgBox("Please confirm your chosen office address: " & interview_location & " Office, " & agency_address.oneline & vbNewLine & vbNewLine & "Press OK to continue, or cancel to end the script." & vbNewLine & vbNewLine & "If this info is incorrect, have an alpha user contact Veronica Cary immediately with the correct address.", vbOKCancel)
' 	If double_check_MsgBox = vbCancel then stopscript
' End if
' 
' 'Checks for MAXIS
' call check_for_MAXIS(False)
' 
' 'Converting the CAF_date variable to a date, for cases where a CAF was turned in
' If no_CAF_check = unchecked then CAF_date = cdate(CAF_date)
' 
' 'Figuring out the last contact day
' If app_type = "recertification" then
'     next_month = datepart("m", dateadd("m", 1, interview_date))
'     next_month_year = datepart("yyyy", dateadd("m", 1, interview_date))
'     last_contact_day = dateadd("d", -1, next_month & "/01/" & next_month_year)
' End if
' If app_type = "new application" then last_contact_day = CAF_date + 30
' If DateDiff("d", interview_date, last_contact_day) < 1 then last_contact_day = interview_date
' 
' 'This checks to make sure the case is not in background and is in the correct footer month for PND1 cases.
' Do
' 	call navigate_to_MAXIS_screen("STAT", "SUMM")
' 	EMReadScreen month_check, 11, 24, 56 'checking for the error message when PND1 cases are not in APPL month
' 	IF left(month_check, 5) = "CASES" THEN 'this means the case can't get into stat in current month
' 		EMWriteScreen mid(month_check, 7, 2), 20, 43 'writing the correct footer month (taken from the error message)
' 		EMWriteScreen mid(month_check, 10, 2), 20, 46 'writing footer year
' 		EMWriteScreen "STAT", 16, 43
' 		EMWriteScreen "SUMM", 21, 70
' 		transmit 'This transmit should take us to STAT / SUMM now
' 	END IF
' 	'This section makes sure the case isn't locked by background, if it is it will loop and try again
' 	EMReadScreen SELF_check, 4, 2, 50
' 	If SELF_check = "SELF" then
' 		PF3
' 		Pause 2
' 	End if
' Loop until SELF_check <> "SELF"
' 
' 'Navigating to SPEC/MEMO
' call navigate_to_MAXIS_screen("SPEC", "MEMO")
' 
' 'Creates a new MEMO. If it's unable the script will stop.
' PF5
' EMReadScreen memo_display_check, 12, 2, 33
' If memo_display_check = "Memo Display" then script_end_procedure("You are not able to go into update mode. Did you enter in inquiry by mistake? Please try again in production.")
' 
' 'Checking for an AREP. If there's an AREP it'll navigate to STAT/AREP, check to see if the forms go to the AREP. If they do, it'll write X's in those fields below.
' row = 4                             'Defining row and col for the search feature.
' col = 1
' EMSearch "ALTREP", row, col         'Row and col are variables which change from their above declarations if "ALTREP" string is found.
' IF row > 4 THEN                     'If it isn't 4, that means it was found.
'     arep_row = row                                          'Logs the row it found the ALTREP string as arep_row
'     call navigate_to_MAXIS_screen("STAT", "AREP")           'Navigates to STAT/AREP to check and see if forms go to the AREP
'     EMReadscreen forms_to_arep, 1, 10, 45                   'Reads for the "Forms to AREP?" Y/N response on the panel.
'     call navigate_to_MAXIS_screen("SPEC", "MEMO")           'Navigates back to SPEC/MEMO
'     PF5                                                     'PF5s again to initiate the new memo process
' END IF
' 'Checking for SWKR
' row = 4                             'Defining row and col for the search feature.
' col = 1
' EMSearch "SOCWKR", row, col         'Row and col are variables which change from their above declarations if "SOCWKR" string is found.
' IF row > 4 THEN                     'If it isn't 4, that means it was found.
'     swkr_row = row                                          'Logs the row it found the SOCWKR string as swkr_row
'     call navigate_to_MAXIS_screen("STAT", "SWKR")         'Navigates to STAT/SWKR to check and see if forms go to the SWKR
'     EMReadscreen forms_to_swkr, 1, 15, 63                'Reads for the "Forms to SWKR?" Y/N response on the panel.
'     call navigate_to_MAXIS_screen("SPEC", "MEMO")         'Navigates back to SPEC/MEMO
'     PF5                                           'PF5s again to initiate the new memo process
' END IF
' EMWriteScreen "x", 5, 10                                        'Initiates new memo to client
' IF forms_to_arep = "Y" THEN EMWriteScreen "x", arep_row, 10     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
' IF forms_to_swkr = "Y" THEN EMWriteScreen "x", swkr_row, 10     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
' transmit                                                        'Transmits to start the memo writing process
' 
' 'Writes the MEMO.
' call write_variable_in_SPEC_MEMO("***********************************************************")
' IF app_type = "new application" then
'     call write_variable_in_SPEC_MEMO("You recently applied for assistance in " & county_name & " on " & CAF_date & ". An interview is required to process your application.")
' Elseif app_type = "recertification" then
'     If no_CAF_check = unchecked then
'         call write_variable_in_SPEC_MEMO("You sent recertification paperwork to " & county_name & " on " & CAF_date & ". An interview is required to process your application.")
'     Else
'         call write_variable_in_SPEC_MEMO("You asked us to set up an interview for your recertification. Remember to send in your forms before the interview.")
'     End if
' End if
' call write_variable_in_SPEC_MEMO("")
' If interview_location = "PHONE" then    'Phone interviews have a different verbiage than any other interview type
' 	IF worker_county_code = "x127" then
' 		call write_variable_in_SPEC_MEMO("Your phone interview is scheduled for " & interview_date & " anytime between " & interview_time & "." )
' 	Else
'     	call write_variable_in_SPEC_MEMO("Your phone interview is scheduled for " & interview_date & " at " & interview_time & "." )
' 	END IF
' Else
'     call write_variable_in_SPEC_MEMO("Your in-office interview is scheduled for " & interview_date & " at " & interview_time & ".")
' End if
' call write_variable_in_SPEC_MEMO("")
' If interview_location = "PHONE" then
' 	if worker_county_code = "x127" then 	'This is for Hennepin County only, x127 recipients/applicants will be calling into the agency using the EZ info number
' 		Call write_variable_in_SPEC_MEMO("Please call the EZ Info Line at 612-596-1300 to complete your phone interview.")
' 		call write_variable_in_SPEC_MEMO("If this date and/or time frame does not work, or you would prefer an interview in the office, please call the EZ Info Line.")
' 	Else
' 		call write_variable_in_SPEC_MEMO("We will be calling you at this number: " & client_phone & ".")
' 		call write_variable_in_SPEC_MEMO("")
'     	call write_variable_in_SPEC_MEMO("If this date and/or time does not work, or you would prefer an interview in the office, please call your worker.")
' 	END IF
' Else
'     call write_variable_in_SPEC_MEMO("Your interview is at the " & interview_location & " Office, located at:")
'     for each line in agency_address.twolines		'"twolines" is an array, so this will write each line in.
' 		call write_variable_in_SPEC_MEMO("   " & line)
'     next
'     call write_variable_in_SPEC_MEMO("")
'     call write_variable_in_SPEC_MEMO("If this date and/or time does not work, or you would prefer an interview over the phone, please call your worker and provide your phone number.")
' End if
' call write_variable_in_SPEC_MEMO("")
' IF app_type = "new application" then            '"deny your application" vs "your case will auto-close"
'     call write_variable_in_SPEC_MEMO("If we do not hear from you by " & last_contact_day & " we will deny your application.")
' Elseif app_type = "recertification" then
'     call write_variable_in_SPEC_MEMO("If we do not hear from you by " & last_contact_day & ", your case will auto-close.")
' END IF
' call write_variable_in_SPEC_MEMO("***********************************************************")
' 
' 'Exits the MEMO
' PF4
' 
' 'Created new variable for TIKL
' interview_info = interview_date & " " & interview_time
' 
' 'TIKLing to remind the worker to send NOMI if appointment is missed
' CALL navigate_to_MAXIS_screen("DAIL", "WRIT")
' CALL create_MAXIS_friendly_date(interview_date, 0, 5, 18)
' Call write_variable_in_TIKL("~*~*~CLIENT WAS SENT AN APPT LETTER FOR INTERVIEW ON " & interview_info & ". IF MISSED SEND NOMI.")
' transmit
' PF3
' 
' 'Navigates to CASE/NOTE and starts a blank one
' start_a_blank_CASE_NOTE
' 
' 'Writes the case note--------------------------------------------
' 'If it's rescheduled, that header should engage. Otherwise, it uses separate headers for new apps and recerts.
' If reschedule_check = checked then
'     call write_variable_in_CASE_NOTE("**Client requested rescheduled appointment, appt letter sent in MEMO**")
' ElseIf app_type = "new application" then
'     call write_variable_in_CASE_NOTE("**New CAF received " & CAF_date & ", appt letter sent in MEMO**")
' ElseIf app_type = "recertification" then
'     If no_CAF_check = unchecked then        'Uses separate headers for whether-or-not a CAF was received.
'         call write_variable_in_CASE_NOTE("**Recert CAF received " & CAF_date & ", appt letter sent in MEMO**")
'     Else
'         call write_variable_in_CASE_NOTE("**Client requested recert appointment, letter sent in MEMO**")
'     End if
' End if
' 
' 'And the rest...
' If same_day_declined_check = checked then write_variable_in_CASE_NOTE("* Same day interview offered and declined.")
' call write_bullet_and_variable_in_CASE_NOTE("Appointment date", interview_date)
' IF interview_location = "PHONE" then
' 	If worker_county_code = "x127" then 	'text for case note for x127 users
' 		call write_bullet_and_variable_in_CASE_NOTE("Appointment time frame", interview_time)
' 		call write_variable_in_CASE_NOTE("* Client was instructed to call the EZ info line to complete interview.")
' 	Else
' 		call write_bullet_and_variable_in_CASE_NOTE("Appointment time", interview_time)
' 		call write_variable_in_CASE_NOTE("* Interview will take place by telephone.")
' 	End if
' Else
' 	call write_bullet_and_variable_in_CASE_NOTE("Appointment time", interview_time)
' 	call write_bullet_and_variable_in_CASE_NOTE("Appointment location", interview_location)
' End if
' call write_bullet_and_variable_in_CASE_NOTE("Why interview is more than six days from now", expedited_explanation)
' call write_bullet_and_variable_in_CASE_NOTE("Client phone", client_phone)
' call write_variable_in_CASE_NOTE("* Client must complete interview by " & last_contact_day & ".")
' IF worker_county_code = "x127" then
' 	call write_variable_in_CASE_NOTE("* TIKL created to call client on interview date. If applicant did not call in, then send NOMI if appropriate.")
' Else
' 	call write_variable_in_CASE_NOTE("* TIKL created for interview date.")
' End if
' If voicemail_check = checked then call write_variable_in_CASE_NOTE("* Left client a voicemail requesting a call back.")
' If forms_to_arep = "Y" then call write_variable_in_CASE_NOTE("* Copy of notice sent to AREP.")              'Defined above
' If forms_to_swkr = "Y" then call write_variable_in_CASE_NOTE("* Copy of notice sent to Social Worker.")     'Defined above
' call write_variable_in_CASE_NOTE("---")
' call write_variable_in_CASE_NOTE(worker_signature)
' 
' script_end_procedure("")
' ===================================================================================================
' ===================================================================================================
' THIS IS THE NEW ONE...
' ===================================================================================================
' ===================================================================================================
'Required for statistical purposes==========================================================================================
name_of_script = "NOTICES - APPOINTMENT LETTER.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 195                               'manual run time in seconds
STATS_denomination = "C"       'C is for each CASE
'END OF stats block=========================================================================================================


'The following defines a class called "address" which carries several simple address properties, which can be used by scripts.
class address
    public street           'Defines a "street" property.
    public city             'Defines a "city" property.
    public state            'Defines a "state" property.
    public zip              'Defines a "zip" property.

    'Creates a "oneline" property containing the entire address on a single line.
    public property get oneline
        oneline = street & ", " & city & ", " & state & " " & zip
    end property

    'Creates a "twolines" property containing the entire address on two lines, split into an array.
    public property get twolines
        twolines = array(street, city & ", " & state & " " & zip)
    end property
end class


'Declaring variables needed by the script
'First, determining the county code. If it isn't declared, it will ask (proxy)

get_county_code

' 10/04/2021 - Removing no longer used office locations.
' agency_office_array = array("Anoka", "Blaine", "Columbia Heights", "Lexington")
' county_office_list = ""     'Blanking this out because it may contain old info from the old global variables (from before this was integrated in this script)
' call convert_array_to_droplist_items(agency_office_array, county_office_list)

'DIALOGS----------------------------------------------------------------------------------------------------
'NOTE: this dialog contains a special modification to allow dynamic creation of the county office list. You cannot edit it in
'   Dialog Editor without modifying the commented line.
BeginDialog appointment_letter_dialog, 0, 0, 156, 355, "Appointment letter"
  EditBox 75, 5, 50, 15, MAXIS_case_number
  DropListBox 50, 25, 95, 15, "new application"+chr(9)+"recertification", app_type
  CheckBox 10, 43, 150, 10, "Check here if this is a reschedule.", reschedule_check
  EditBox 50, 55, 95, 15, CAF_date
  CheckBox 10, 75, 130, 10, "Check here if this is a recert and the", no_CAF_check
  DropListBox 70, 100, 75, 15, "Select one..."+chr(9)+"PHONE"+chr(9)+"Blaine", interview_location     'This line dynamically creates itself based on the information in the FUNCTIONS FILE.
  EditBox 70, 120, 75, 15, interview_date
  EditBox 70, 140, 75, 15, interview_time
  EditBox 80, 160, 65, 15, client_phone
  CheckBox 10, 200, 95, 10, "Client appears expedited", expedited_check
  CheckBox 10, 215, 135, 10, "Same day interview offered/declined", same_day_declined_check
  EditBox 10, 250, 135, 15, expedited_explanation
  CheckBox 10, 280, 135, 10, "Check here if you left V/M with client", voicemail_check
  EditBox 85, 305, 60, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 25, 325, 50, 15
    CancelButton 85, 325, 50, 15
  Text 25, 10, 50, 10, "Case number:"
  Text 15, 30, 30, 10, "App type:"
  Text 15, 60, 35, 10, "CAF date:"
  Text 30, 85, 105, 10, "CAF hasn't been received yet."
  Text 15, 105, 55, 10, "Int'vw location:"
  Text 15, 125, 50, 10, "Interview date: "
  Text 15, 145, 50, 10, "Interview time:"
  Text 15, 160, 60, 20, "Client phone (if phone interview):"
  GroupBox 5, 185, 145, 85, "Expedited questions"
  Text 10, 230, 135, 20, "If expedited interview date is not within six days of the application, explain:"
  Text 45, 290, 75, 10, "requesting a call back."
  Text 15, 310, 65, 10, "Worker signature:"
EndDialog

'Case number only dialog for x127 users
BeginDialog case_number_dialog, 0, 0, 136, 60, "Case number dialog"
  EditBox 60, 10, 60, 15, MAXIS_case_number
  ButtonGroup ButtonPressed
    OkButton 15, 35, 50, 15
    CancelButton 70, 35, 50, 15
  Text 10, 15, 45, 10, "Case number:"
EndDialog

'Hennepin County appointment letter
BeginDialog Hennepin_appt_dialog, 0, 0, 296, 75, "Hennepin County appointment letter"
  EditBox 205, 25, 55, 15, interview_date
  EditBox 65, 50, 115, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 185, 50, 50, 15
    CancelButton 240, 50, 50, 15
  EditBox 65, 25, 55, 15, CAF_date
  Text 5, 55, 60, 10, "Worker signature:"
  Text 140, 30, 60, 10, "Appointment date:"
  GroupBox 20, 10, 255, 35, "Enter a new appointment date only if it's a date county offices are not open."
  Text 30, 30, 35, 10, "CAF date:"
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connects to BlueZone & gathers case number
EMConnect ""
call MAXIS_case_number_finder(MAXIS_case_number)

Do
	Do
		err_msg = ""
		Dialog appointment_letter_dialog
		If ButtonPressed = cancel then stopscript
		If isnumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & vbnewline & "* You must fill in a valid case number. Please try again."
		CAF_date = replace(CAF_date, ".", "/")
		If no_CAF_check = checked and app_type = "new application" then no_CAF_check = unchecked 'Shuts down "no_CAF_check" so that it will validate the date entered. New applications can't happen if a CAF wasn't provided.
		If no_CAF_check = unchecked and isdate(CAF_date) = False then err_msg = err_msg & vbnewline & "* You did not enter a valid CAF date (MM/DD/YYYY format). Please try again."
    	If interview_location = "Select one..." then err_msg = err_msg & vbnewline & "* You must select an interview location. Please try again."
    	If interview_location = "PHONE" and client_phone = "" then err_msg = err_msg & vbnewline & "* If this is a phone interview, you must enter a phone number. Please try again."
    	interview_date = replace(interview_date, ".", "/")
    	If isdate(interview_date) = False then err_msg = err_msg & vbnewline & "* You did not enter a valid interview date (MM/DD/YYYY format). Please try again."
    	If interview_time = "" then err_msg = err_msg & vbnewline & "* You must type an interview time. Please try again."
    	If no_CAF_check = checked then exit do 'If no CAF was turned in, this layer of validation is unnecessary, so the script will skip it.
    	If expedited_check = checked and datediff("d", CAF_date, interview_date) > 6 and expedited_explanation = "" then err_msg = err_msg & vbnewline & "* You have indicated that your case is expedited, but scheduled the interview date outside of the six-day window. An explanation of why is required for QC purposes."
		If worker_signature = "" then err_msg = err_msg & vbnewline & "* You must provide a signature for your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	Loop until err_msg = ""
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

'Creates a variable to contain the agency addresses. "Address" is a class defined above.
set agency_address = new address

'As these are all MN intake locations, the state for all of them will be MN.
agency_address.state = "MN"


' 10/04/2021 - Updating address information following site move.
agency_address.street = "1201 89th Ave, Suite 400"
agency_address.city = "Blaine"
agency_address.zip = "55434"

' IF interview_location = "Anoka" THEN
'     agency_address.street = "2100 3rd Ave, Suite 400"
'     agency_address.city = "Anoka"
'     agency_address.zip = "55303"
' ELSEIF interview_location = "Blaine" THEN
'     agency_address.street = "1201 89th Ave, Suite 400"
'     agency_address.city = "Blaine"
'     agency_address.zip = "55434"
' ELSEIF interview_location = "Columbia Heights" THEN
'     agency_address.street = "3980 Central Ave NE"
'     agency_address.city = "Columbia Heights"
'     agency_address.zip = "55421"
' ELSEIF interview_location = "Lexington" THEN
'     agency_address.street = "9201 S. Highway Drive Suite B"
'     agency_address.city = "Lexington"
'     agency_address.zip = "55014"
' END IF

'This is a temporary MsgBox that expires 09/01/2015. It is designed to "make sure" that the address is correct. Because this function is new, I want to be ABSOLUTELY SURE it's working before notices get sent out.
If interview_location <> "PHONE" and datediff("d", date, #9/1/2015#) > 0 then
	double_check_MsgBox = MsgBox("Please confirm your chosen office address: " & interview_location & " Office, " & agency_address.oneline & vbNewLine & vbNewLine & "Press OK to continue, or cancel to end the script." & vbNewLine & vbNewLine & "If this info is incorrect, have an alpha user contact Veronica Cary immediately with the correct address.", vbOKCancel)
	If double_check_MsgBox = vbCancel then stopscript
End if

'Checks for MAXIS
call check_for_MAXIS(False)

'Converting the CAF_date variable to a date, for cases where a CAF was turned in
If no_CAF_check = unchecked then CAF_date = cdate(CAF_date)

'Figuring out the last contact day
If app_type = "recertification" then
    next_month = datepart("m", dateadd("m", 1, interview_date))
    next_month_year = datepart("yyyy", dateadd("m", 1, interview_date))
    last_contact_day = dateadd("d", -1, next_month & "/01/" & next_month_year)
End if
If app_type = "new application" then last_contact_day = CAF_date + 29
If DateDiff("d", interview_date, last_contact_day) < 1 then last_contact_day = interview_date

'This checks to make sure the case is not in background and is in the correct footer month for PND1 cases.
Do
	call navigate_to_MAXIS_screen("STAT", "SUMM")
	EMReadScreen month_check, 11, 24, 56 'checking for the error message when PND1 cases are not in APPL month
	IF left(month_check, 5) = "CASES" THEN 'this means the case can't get into stat in current month
		EMWriteScreen mid(month_check, 7, 2), 20, 43 'writing the correct footer month (taken from the error message)
		EMWriteScreen mid(month_check, 10, 2), 20, 46 'writing footer year
		EMWriteScreen "STAT", 16, 43
		EMWriteScreen "SUMM", 21, 70
		transmit 'This transmit should take us to STAT / SUMM now
	END IF
	'This section makes sure the case isn't locked by background, if it is it will loop and try again
	EMReadScreen SELF_check, 4, 2, 50
	If SELF_check = "SELF" then
		PF3
		Pause 2
	End if
Loop until SELF_check <> "SELF"

'Navigating to SPEC/MEMO and starting a new memo
start_a_new_spec_memo

'Writes the MEMO.
call write_variable_in_SPEC_MEMO("***********************************************************")
IF app_type = "new application" then
    call write_variable_in_SPEC_MEMO("You recently applied for assistance in " & county_name & " on " & CAF_date & ". An interview is required to process your application.")
Elseif app_type = "recertification" then
    If no_CAF_check = unchecked then
        call write_variable_in_SPEC_MEMO("You sent recertification paperwork to " & county_name & " on " & CAF_date & ". An interview is required to process your application.")
    Else
        call write_variable_in_SPEC_MEMO("You asked us to set up an interview for your recertification. Remember to send in your forms before the interview.")
    End if
End if
call write_variable_in_SPEC_MEMO("")
If interview_location = "PHONE" then    'Phone interviews have a different verbiage than any other interview type
	IF worker_county_code = "x127" then
		call write_variable_in_SPEC_MEMO("Your phone interview is scheduled for " & interview_date & " anytime between " & interview_time & "." )
	Else
    	call write_variable_in_SPEC_MEMO("Your phone interview is scheduled for " & interview_date & " at " & interview_time & "." )
	END IF
Else
    call write_variable_in_SPEC_MEMO("Your in-office interview is scheduled for " & interview_date & " at " & interview_time & ".")
End if
call write_variable_in_SPEC_MEMO("")
If interview_location = "PHONE" then
	if worker_county_code = "x127" then 	'This is for Hennepin County only, x127 recipients/applicants will be calling into the agency using the EZ info number
		Call write_variable_in_SPEC_MEMO("Please call the EZ Info Line at 612-596-1300 to complete your phone interview.")
		call write_variable_in_SPEC_MEMO("If this date and/or time frame does not work, or you would prefer an interview in the office, please call the EZ Info Line.")
	Else
		call write_variable_in_SPEC_MEMO("We will be calling you at this number: " & client_phone & ".")
		call write_variable_in_SPEC_MEMO("")
    	call write_variable_in_SPEC_MEMO("If this date and/or time does not work, or you would prefer an interview in the office, please call your worker.")
	END IF
Else
    call write_variable_in_SPEC_MEMO("Your interview is at the " & interview_location & " Office, located at:")
    for each line in agency_address.twolines		'"twolines" is an array, so this will write each line in.
		call write_variable_in_SPEC_MEMO("   " & line)
    next
    call write_variable_in_SPEC_MEMO("")
    call write_variable_in_SPEC_MEMO("If this date and/or time does not work, or you would prefer an interview over the phone, please call your worker and provide your phone number.")
End if
call write_variable_in_SPEC_MEMO("")
IF app_type = "new application" then            '"deny your application" vs "your case will auto-close"
    call write_variable_in_SPEC_MEMO("If we do not hear from you by " & last_contact_day & " we will deny your application.")
Elseif app_type = "recertification" then
    call write_variable_in_SPEC_MEMO("If we do not hear from you by " & last_contact_day & ", your case will auto-close.")
END IF
call write_variable_in_SPEC_MEMO("***********************************************************")

'Exits the MEMO
PF4

'Created new variable for TIKL
interview_info = interview_date & " " & interview_time

'TIKLing to remind the worker to send NOMI if appointment is missed
CALL navigate_to_MAXIS_screen("DAIL", "WRIT")
CALL create_MAXIS_friendly_date(interview_date, 0, 5, 18)
Call write_variable_in_TIKL("~*~*~CLIENT WAS SENT AN APPT LETTER FOR INTERVIEW ON " & interview_info & ". IF MISSED SEND NOMI.")
transmit
PF3

'Navigates to CASE/NOTE and starts a blank one
start_a_blank_CASE_NOTE

'Writes the case note--------------------------------------------
'If it's rescheduled, that header should engage. Otherwise, it uses separate headers for new apps and recerts.
If reschedule_check = checked then
    call write_variable_in_CASE_NOTE("**Client requested rescheduled appointment, appt letter sent in MEMO**")
ElseIf app_type = "new application" then
    call write_variable_in_CASE_NOTE("**New CAF received " & CAF_date & ", appt letter sent in MEMO**")
ElseIf app_type = "recertification" then
    If no_CAF_check = unchecked then        'Uses separate headers for whether-or-not a CAF was received.
        call write_variable_in_CASE_NOTE("**Recert CAF received " & CAF_date & ", appt letter sent in MEMO**")
    Else
        call write_variable_in_CASE_NOTE("**Client requested recert appointment, letter sent in MEMO**")
    End if
End if

'And the rest...
If same_day_declined_check = checked then write_variable_in_CASE_NOTE("* Same day interview offered and declined.")
call write_bullet_and_variable_in_CASE_NOTE("Appointment date", interview_date)
IF interview_location = "PHONE" then
	If worker_county_code = "x127" then 	'text for case note for x127 users
		call write_bullet_and_variable_in_CASE_NOTE("Appointment time frame", interview_time)
		call write_variable_in_CASE_NOTE("* Client was instructed to call the EZ info line to complete interview.")
	Else
		call write_bullet_and_variable_in_CASE_NOTE("Appointment time", interview_time)
		call write_variable_in_CASE_NOTE("* Interview will take place by telephone.")
	End if
Else
	call write_bullet_and_variable_in_CASE_NOTE("Appointment time", interview_time)
	call write_bullet_and_variable_in_CASE_NOTE("Appointment location", interview_location)
End if
call write_bullet_and_variable_in_CASE_NOTE("Why interview is more than six days from now", expedited_explanation)
call write_bullet_and_variable_in_CASE_NOTE("Client phone", client_phone)
call write_variable_in_CASE_NOTE("* Client must complete interview by " & last_contact_day & ".")
IF worker_county_code = "x127" then
	call write_variable_in_CASE_NOTE("* TIKL created to call client on interview date. If applicant did not call in, then send NOMI if appropriate.")
Else
	call write_variable_in_CASE_NOTE("* TIKL created for interview date.")
End if
If voicemail_check = checked then call write_variable_in_CASE_NOTE("* Left client a voicemail requesting a call back.")
If forms_to_arep = "Y" then call write_variable_in_CASE_NOTE("* Copy of notice sent to AREP.")              'Defined above
If forms_to_swkr = "Y" then call write_variable_in_CASE_NOTE("* Copy of notice sent to Social Worker.")     'Defined above
call write_variable_in_CASE_NOTE("---")
call write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure("")
