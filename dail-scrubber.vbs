'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'    HOW THE DAIL SCRUBER WORKS:
'
'    This script opens up other script files, using a custom function (run_DAIL_scrubber_script), followed by the path to the script file. It's done this
'      way because there could be hundreds of DAIL messages, and to work all of the combinations into one script would be incredibly tedious and long.
'
'    This script works by moving the message (where the cursor is located) to the top of the screen, and then reading the message text. Whatever the
'      message text says dictates which script loads up.
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

' 'Required for statistical purposes===============================================================================
' name_of_script = "DAIL - DAIL SCRUBBER.vbs"
' start_time = timer
' 
' 'Because we are running these locally, we are going to get rid of all the calls to GitHub...
' if func_lib_run <> true then 
' 	FuncLib_URL = "I:\Blue Zone Scripts\Functions Library.vbs"
' 	Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
' 	Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
' 	text_from_the_other_script = fso_command.ReadAll
' 	fso_command.Close
' 	Execute text_from_the_other_script
' 	func_lib_run = true
' end if
'END FUNCTIONS LIBRARY BLOCK================================================================================================

Set objNet = CreateObject("WScript.NetWork") 
windows_user_ID = UCASE(objNet.UserName)

IF windows_user_ID = "SCOSTENS" THEN 
  Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
  Set fso_command = run_another_script_fso.OpenTextFile("I:\Blue Zone Scripts\Public Assistance Script Files\DHS-MAXIS-Scripts-master\dail\dail-scrubber-test.vbs")
  text_from_the_other_script = fso_command.ReadAll
  fso_command.Close
  ExecuteGlobal text_from_the_other_script
END IF

script_repository = "I:\Blue Zone Scripts\Public Assistance Script Files\DHS-MAXIS-Scripts-master\"

FUNCTION launch_selected_script(script_file_path)
  Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
  Set fso_command = run_another_script_fso.OpenTextFile(script_file_path)
  text_from_the_other_script = fso_command.ReadAll
  fso_command.Close
  ExecuteGlobal text_from_the_other_script
  func_lib_run = true
END FUNCTION

FUNCTION transmit
	EMSendKey "<Enter>"
	EMWaitReady 0,0
END FUNCTION

' THIS FUNCTIONALITY MUST BE COMMENTED-OUT...THE SCRIPT CLASS IS ALREADY DECLARED IN THE FUNCTIONS LIBRARY.
' 'A class for each script item
' class script
' 
' 	public script_name             	'The familiar name of the script
' 	public file_name               	'The actual file name
' 	public description             	'The description of the script
' 	public button                  	'A variable to store the actual results of ButtonPressed (used by much of the script functionality)
'     public category               	'The script category (ACTIONS/BULK/etc)
'     public SIR_instructions_URL    	'The instructions URL in SIR
'     public button_plus_increment	'Workflow scripts use a special increment for buttons (adding or subtracting from total times to run). This is the add button.
' 	public button_minus_increment	'Workflow scripts use a special increment for buttons (adding or subtracting from total times to run). This is the minus button.
' 	public total_times_to_run		'A variable for the total times the script should run
' 	public subcategory				'An array of all subcategories a script might exist in, such as "LTC" or "A-F"
' 
' 	public property get button_size	'This part determines the size of the button dynamically by determining the length of the script name, multiplying that by 3.5, rounding the decimal off, and adding 10 px
' 		button_size = round ( len( script_name ) * 3.5 ) + 10
' 	end property
' 
' end class

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
' changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
' call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
' changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'CONNECTS TO DEFAULT SCREEN
EMConnect ""

'CHECKS TO MAKE SURE THE WORKER IS ON THEIR DAIL
EMReadscreen dail_check, 4, 2, 48
If dail_check <> "DAIL" then script_end_procedure("You are not in your dail. This script will stop.")

'TYPES A "T" TO BRING THE SELECTED MESSAGE TO THE TOP
EMSendKey "t"
transmit

'The following reads the message in full for the end part (which tells the worker which message was selected)
EMReadScreen full_message, 58, 6, 20

'THE FOLLOWING CODES ARE THE INDIVIDUAL MESSAGES. IT READS THE MESSAGE, THEN CALLS A NEW SCRIPT.----------------------------------------------------------------------------------------------------

'Random messages generated from an affiliated case (loads AFFILIATED CASE LOOKUP) OR XFS Closed for Postponed Verifications (loads POSTPONTED XFS VERIFICATIONS)
'Both of these messages start with 'FS' on the DAIL, so they need to be nested, or it never gets passed the affilated case look up
EMReadScreen stat_check, 4, 6, 6
If stat_check = "FS  " or stat_check = "HC  " or stat_check = "GA  " or stat_check = "MSA " or stat_check = "STAT" then
	'now it checks if you are acctually running from a XFS Autoclosed DAIL. These messages don't have an affiliated case attached - so there will be no overlap
	EMReadScreen xfs_check, 49, 6, 20
	If xfs_check = "CASE AUTO-CLOSED FOR FAILURE TO PROVIDE POSTPONED" then
		call launch_selected_script(script_repository & "dail\postponed-expedited-snap-verifications.vbs")
	Else
		call launch_selected_script(script_repository & "dail\affiliated-case-lookup.vbs")
	End If
End If

'Checking for 12 month contact TIKL from CAF and CAR scripts(loads NOTICES - 12 month contact)
EMReadScreen twelve_mo_contact_check, 57, 6, 20
IF twelve_mo_contact_check = "IF SNAP IS OPEN, REVIEW TO SEE IF 12 MONTH CONTACT LETTER" THEN
	EMReadScreen MAXIS_case_number, 8, 5, 73									'reading the case number for ease of use
	MAXIS_case_number = TRIM(MAXIS_case_number)							'trimming the blank spaces
	func_lib_run = true
	launch_selected_script(script_repository & "notices\12-month-contact.vbs")
END IF

'RSDI/BENDEX info received by agency (loads BNDX SCRUBBER)
EMReadScreen BENDEX_check, 47, 6, 30
If BENDEX_check = "BENDEX INFORMATION HAS BEEN STORED - CHECK INFC" then 

	BeginDialog delete_message_dialog, 0, 0, 126, 45, "Double-Check the Computer's Work..."
	ButtonGroup ButtonPressed
		PushButton 10, 25, 50, 15, "YES", delete_button
		PushButton 60, 25, 50, 15, "NO", do_not_delete
	Text 30, 10, 65, 10, "Delete the DAIL??"
	EndDialog
	
	'------------------THIS SCRIPT IS DESIGNED TO BE RUN FROM THE DAIL SCRUBBER.
	'------------------As such, it does NOT include protections to be ran independently.
	
	EMConnect ""
	
	EMReadScreen on_dail, 4, 2, 48
	IF on_dail <> "DAIL" THEN script_end_procedure("You are not in DAIL. Please navigate to DAIL and run the script again.")
	
	EMGetCursor read_row, read_column
	
	EMReadScreen is_right_line, 34, read_row, 30
	IF is_right_line <> "BENDEX INFORMATION HAS BEEN STORED" THEN script_end_procedure("You are not on the correct line. Please select a BNDX message on your DAIL.")
	EMReadScreen original_bndx_dail, 30, read_row, 6
	
	EMReadScreen cl_ssn, 9, read_row, 20
		ssn_first = left(cl_ssn, 3)
		ssn_first = ssn_first & " "
		ssn_mid = right(left(cl_ssn, 5), 2)
		ssn_mid = ssn_mid & " "
		ssn_end = right(cl_ssn, 4)
		use_ssn = ssn_first & ssn_mid & ssn_end
	search_row = read_row
	
	'========== Collects the case number ==========
	DO
		EMReadScreen look_for_case_number, 18, search_row, 63
		IF left(look_for_case_number, 10) = "CASE NBR: " THEN
			maxis_case_number = right(look_for_case_number, 8)
			maxis_case_number = replace(maxis_case_number, " ", "")
		ELSE
			search_row = search_row - 1
		END IF
	LOOP UNTIL left(look_for_case_number, 10) = "CASE NBR: "
	
	EMWriteScreen "I", read_row, 3
	transmit
	EMWriteScreen "BNDX", 20, 71
	transmit
	
	'checking for IRS non-disclosure agreement.
	EMReadScreen agreement_check, 9, 2, 24
	IF agreement_check = "Automated" THEN script_end_procedure("To view INFC data you will need to review the agreement. Please navigate to INFC and then into one of the screens and review the agreement.")
	
	DIM bndx_array()
	ReDim bndx_array(2, 5)
	'========== Collects information from BNDX ==========
	EMReadScreen bndx_claim_one_number, 13, 5, 12
	bndx_claim_one_number = replace(bndx_claim_one_number, " ", "")
	EMReadScreen bndx_claim_one_amt, 8, 7, 12
	bndx_claim_one_amt = replace(bndx_claim_one_amt, " ", "")
	'ReDim bndx_array(0, 5)
	num_of_rsdi = 0
	bndx_array(0, 0) = bndx_claim_one_number
	bndx_array(0, 1) = bndx_claim_one_amt
	
	EMReadScreen bndx_claim_two_number, 13, 5, 38
	bndx_claim_two_number = replace(bndx_claim_two_number, " ", "")
	EMReadScreen bndx_claim_two_amt, 8, 7, 38
	bndx_claim_two_amt = replace(bndx_claim_two_amt, " ", "")
		IF bndx_claim_two_amt <> "" THEN
	'		ReDim bndx_array(1, 5)
			num_of_rsdi = 1
			bndx_array(1, 0) = bndx_claim_two_number
			bndx_array(1, 1) = bndx_claim_two_amt
		END IF
	
	EMReadScreen bndx_claim_three_number, 13, 5, 64
	bndx_claim_three_number = replace(bndx_claim_three_number, " ", "")
	EMReadScreen bndx_claim_three_amt, 8, 7, 64
	bndx_claim_three_amt = replace(bndx_claim_three_amt, " ", "")
		IF bndx_claim_three_amt <> "" THEN
	'		ReDim bndx_array(2, 5)
			num_of_rsdi = 2
			bndx_array(2, 0) = bndx_claim_three_number
			bndx_array(2, 1) = bndx_claim_three_amt
		END IF
	
	
	
	'========== Goes back to STAT/PROG to determine which programs are active. ==========
	back_to_SELF
	EMWriteScreen "STAT", 16, 43
	EMWriteScreen maxis_case_number, 18, 43
	EMWriteScreen "PROG", 21, 70
	transmit
	EMReadScreen abended_check, 7, 9, 27
	IF abended_check = "abended" THEN transmit
	EMReadScreen errr_check, 4, 2, 52
	IF errr_check = "ERRR" THEN transmit
	
	EMReadScreen cash_one_status, 4, 6, 74
	EMReadScreen cash_two_status, 4, 7, 74
	EMReadScreen grh_status, 4, 9, 74
	EMReadScreen fs_status, 4, 10, 74
	EMReadScreen ive_status, 4, 11, 74
	EMReadScreen hc_status, 4, 12, 74
	
	IF cash_one_status <> "ACTV" AND cash_two_status <> "ACTV" AND grh_status <> "ACTV" AND fs_status <> "ACTV" AND ive_status <> "ACTV" AND hc_status <> "ACTV" THEN
	IF cash_one_status <> "PEND" AND cash_two_status <> "PEND" AND grh_status <> "PEND" AND fs_status <> "PEND" AND ive_status <> "PEND" AND hc_status <> "PEND" THEN script_end_procedure("The client does not have any active or pending MAXIS cases.")
	END IF
	
	'========== Navigates to MEMB to grab the member number for cases in which there are mulitple persons on the case with a BNDX message. ==========
	EMWriteScreen "MEMB", 20, 71
	transmit
	
	DO
		EMReadScreen memb_ssn, 11, 7, 42
		IF use_ssn = memb_ssn THEN
			EMReadScreen reference_number, 2, 4, 33
		ELSE
			transmit
		END IF
	LOOP UNTIL use_ssn = memb_ssn
	
	FOR i = 0 TO num_of_rsdi
		end_of_unea = ""
		'========== Goes to STAT/UNEA ==========
		EMWriteScreen "UNEA", 20, 71
		EMWriteScreen reference_number, 20, 76
		EMWriteScreen "01", 20, 79
		transmit
	
		EMReadScreen number_of_unea_panels, 1, 2, 78
		IF number_of_unea_panels = "0" THEN
			script_end_procedure("Client is not showing any UNEA panels.")
		ELSEIF number_of_unea_panels = "1" THEN
			EMReadScreen unea_type, 4, 5, 40
			IF unea_type = "RSDI" THEN
				EMReadScreen unea_claim_number, 11, 6, 37
				bndx_array(i, 2) = unea_claim_number
				IF (right(bndx_array(i, 0), 1) = "A" AND right(bndx_array(i, 0), 2) <> "HA") OR _
					(right(bndx_array(i, 0), 1) = "B" AND right(bndx_array(i, 0), 2) <> "HB") OR _
					right(bndx_array(i, 0), 1) = "D" OR _
					right(bndx_array(i, 0), 1) = "E" OR _
					right(bndx_array(i, 0), 1) = "G" OR _
					right(bndx_array(i, 0), 1) = "M" OR _
					right(bndx_array(i, 0), 1) = "T" OR _
					right(bndx_array(i, 0), 1) = "W" THEN bndx_array(i, 2) = left(bndx_array(i, 2), 10)
				IF bndx_array(i, 0) <> bndx_array(i, 2) THEN error_message = error_message & chr(13) & "Claim numbers do not match."
				EMReadScreen unea_prospective_amt, 8, 18, 68
				bndx_array(i, 3) = trim(unea_prospective_amt)
				IF ((CDbl(bndx_array(i, 3)) - CDBl(bndx_array(i, 1)) > county_bndx_variance_threshold) OR (CDbl(bndx_array(i, 1)) - CDbl(bndx_array(i, 3)) > county_bndx_variance_threshold)) THEN error_message = error_message & chr(13) & "The prospective amount in UNEA is significantly different from BNDX for BNDX claim " & (i + 1) & ", " & bndx_array(i, 0) & "."
				IF fs_status = "ACTV" or fs_status = "PEND" THEN
					EMWriteScreen "X", 10, 26
					transmit
					EMReadScreen unea_pic_amt, 8, 18, 56
					bndx_array(i, 4) = trim(unea_pic_amt)
					IF ((CDbl(bndx_array(i, 4)) - CDbl(bndx_array(i, 1)) > county_bndx_variance_threshold) OR (CDbl(bndx_array(i, 1)) - CDbl(bndx_array(i, 4)) > county_bndx_variance_threshold)) THEN error_message = error_message & chr(13) & "The claim amount on the PIC is significantly different from BNDX for BNDX claim " & (i + 1) & ", " & bndx_array(i, 0) & "."
					PF3
				ELSE
					bndx_array(i, 4) = ""
				END IF
				IF hc_status = "ACTV" or hc_status = "PEND" THEN
					EMWriteScreen "X", 6, 56
					transmit
					EMReadScreen unea_hc_inc_amt, 8, 9, 65
					bndx_array(i, 5) = trim(unea_hc_inc_amt)
					IF ((CDbl(bndx_array(i, 5)) - CDbl(bndx_array(i, 1)) > county_bndx_variance_threshold) OR (CDbl(bndx_array(i, 1)) - CDbl(bndx_array(i, 5)) > county_bndx_variance_threshold)) THEN error_message = error_message & chr(13) & "The claim amount on the HC Inc Est is significantly different from BNDX for BNDX claim " & (i + 1) & ", " & bndx_array(i, 0) & "."
					PF3
				ELSE
					bndx_array(i, 5) = ""
				END IF
			ELSE
				script_end_procedure("This case is not showing an RSDI claim.")
			END IF
		ELSE
			DO
				EMReadScreen unea_type, 4, 5, 40
				IF unea_type <> "RSDI" THEN transmit
				EMReadScreen end_of_unea, 15, 24, 2
				end_of_unea = trim(end_of_unea)
				IF end_of_unea <> "" THEN error_message = error_message & vbCr & "There is a discrepancy with BNDX claim " & (i + 1) & ", " & bndx_array(i, 0) & "."
			LOOP UNTIL unea_type = "RSDI" or end_of_unea <> ""
			IF end_of_unea = "" THEN
				DO
					EMReadScreen unea_claim_number, 11, 6, 37
					bndx_array(i, 2) = unea_claim_number
					IF (right(bndx_array(i, 0), 1) = "A" AND right(bndx_array(i, 0), 2) <> "HA") OR _
						(right(bndx_array(i, 0), 1) = "B" AND right(bndx_array(i, 0), 2) <> "HB") OR _
						right(bndx_array(i, 0), 1) = "D" OR _
						right(bndx_array(i, 0), 1) = "E" OR _
						right(bndx_array(i, 0), 1) = "G" OR _
						right(bndx_array(i, 0), 1) = "M" OR _
						right(bndx_array(i, 0), 1) = "T" OR _
						right(bndx_array(i, 0), 1) = "W" THEN bndx_array(i, 2) = left(bndx_array(i, 2), 10)
					IF bndx_array(i, 0) <> bndx_array(i, 2) THEN transmit
					EMReadScreen end_of_unea, 15, 24, 2
					end_of_unea = trim(end_of_unea)
					IF end_of_unea <> "" THEN error_message = error_message & vbCr & "There is a discrepancy with BNDX claim " & (i + 1) & ", " & bndx_array(i, 0) & "."
				LOOP UNTIL bndx_array(i, 0) = bndx_array(i, 2) OR end_of_unea <> ""
				IF end_of_unea = "" THEN
					EMReadScreen unea_prospective_amt, 8, 18, 68
					bndx_array(i, 3) = trim(unea_prospective_amt)
					IF ((CDbl(bndx_array(i, 3)) - CDBl(bndx_array(i, 1)) > county_bndx_variance_threshold) OR (CDbl(bndx_array(i, 1)) - CDbl(bndx_array(i, 3)) > county_bndx_variance_threshold)) THEN error_message = error_message & chr(13) & "The prospective amount in UNEA is significantly different from BNDX for BNDX claim " & (i + 1) & ", " & bndx_array(i, 0) & "."
					IF fs_status = "ACTV" or fs_status = "PEND" THEN
						EMWriteScreen "X", 10, 26
						transmit
						EMReadScreen unea_pic_amt, 8, 18, 56
						bndx_array(i, 4) = trim(unea_pic_amt)
						IF ((CDbl(bndx_array(i, 4)) - CDbl(bndx_array(i, 1)) > county_bndx_variance_threshold) OR (CDbl(bndx_array(i, 1)) - CDbl(bndx_array(i, 4)) > county_bndx_variance_threshold)) THEN error_message = error_message & chr(13) & "The claim amount on the PIC is significantly different from BNDX for BNDX claim " & (i + 1) & ", " & bndx_array(i, 0) & "."
						PF3
					ELSE
						bndx_array(i, 4) = ""
					END IF
					IF hc_status = "ACTV" or hc_status = "PEND" THEN
						EMWriteScreen "X", 6, 56
						transmit
						EMReadScreen unea_hc_inc_amt, 8, 9, 65
						bndx_array(i, 5) = trim(unea_hc_inc_amt)
						IF ((CDbl(bndx_array(i, 5)) - CDbl(bndx_array(i, 1)) > county_bndx_variance_threshold) OR (CDbl(bndx_array(i, 1)) - CDBl(bndx_array(i, 5)) > county_bndx_variance_threshold)) THEN error_message = error_message & chr(13) & "The claim amount on the HC Inc Est is significantly different from BNDX for BNDX claim " & (i + 1) & ", " & bndx_array(i, 0) & "."
						PF3
					ELSE
						bndx_array(i, 5) = ""
					END IF
				END IF
			END IF
		END IF
	NEXT
	
	back_to_SELF
	EMWriteScreen "DAIL", 16, 43
	EMWriteScreen maxis_case_number, 18, 43
	EMWriteScreen "DAIL", 21, 70
	transmit
	
	'========== The bit about the MSGBox is used only as a safeguard for Beta Testing.
	IF error_message = "" THEN
		compare_message = "BNDX Conclusion" & vbCr & "============="
		FOR i = 0 to num_of_rsdi
			compare_message = compare_message & vbCr & "BNDX Claim #: " & bndx_array(i, 0)
			compare_message = compare_message & vbCr & "  BNDX Amt: " & bndx_array(i, 1)
			compare_message = compare_message & vbCr & "  UNEA Prosp Amt: " & bndx_array(i, 3)
			IF bndx_array(i, 4) <> "" THEN compare_message = compare_message & vbCr & "  SNAP PIC Amt: " & bndx_array(i, 4)
			IF bndx_array(i, 5) <> "" THEN compare_message = compare_message & vbCr & "  HC Inc Est Amt: " & bndx_array(i, 5)
		NEXT
		MSGBox compare_message
		DIALOG delete_message_dialog
			IF ButtonPressed = delete_button THEN
				DO
					dail_read_row = 6
					DO
						EMReadScreen double_check, 30, dail_read_row, 6
						IF double_check = original_bndx_dail THEN
							EMWriteScreen "D", dail_read_row, 3
							transmit
							EXIT DO
						ELSE
							dail_read_row = dail_read_row + 1
						END IF
						IF dail_read_row = 19 THEN PF8
					LOOP UNTIL dail_read_row = 19
				LOOP UNTIL double_check = original_bndx_dail
			END IF
	ELSE
		error_message = "*** NOTICE ***" & vbCr & "==========" & vbCr & error_message & vbCr & vbCr & "Review case and request RSDI information if necessary."
		MSGBox error_message
		'compare_message = "BNDX Conclusion" & vbCr & "============="
		'FOR i = 0 to num_of_rsdi
		'	compare_message = compare_message & vbCr & "BNDX Claim #: " & bndx_array(i, 0)
		'	compare_message = compare_message & vbCr & "  BNDX Amt: " & bndx_array(i, 1)
		'	compare_message = compare_message & vbCr & "  UNEA Prosp Amt: " & bndx_array(i, 3)
		'	IF bndx_array(i, 4) <> "" THEN compare_message = compare_message & vbCr & "  SNAP PIC Amt: " & bndx_array(i, 4)
		'	IF bndx_array(i, 5) <> "" THEN compare_message = compare_message & vbCr & "  HC Inc Est Amt: " & bndx_array(i, 5)
		'NEXT
		'MSGBox compare_message
	END IF
	
	script_end_procedure("")	
	
	' call launch_selected_script(script_repository & "dail\bndx-scrubber.vbs")
	
END IF

'CIT/ID has been verified through the SSA (loads CITIZENSHIP VERIFIED)
EMReadScreen CIT_check, 46, 6, 20
If CIT_check = "MEMI:CITIZENSHIP HAS BEEN VERIFIED THROUGH SSA" then call launch_selected_script(script_repository & "dail\citizenship-verified.vbs")

'CS reports a new employer to the worker (loads CS REPORTED NEW EMPLOYER)
EMReadScreen CS_new_emp_check, 25, 6, 20
If CS_new_emp_check = "CS REPORTED: NEW EMPLOYER" then call launch_selected_script(script_repository & "dail\cs-reported-new-employer.vbs")

'Child support messages (loads CSES PROCESSING)
EMReadScreen CSES_check, 4, 6, 6
If CSES_check = "CSES" or CSES_check = "TIKL" then		'TIKLs are used for fake cases and testing
	EMReadScreen CSES_DISB_check, 4, 6, 20				'Checks for the DISB string, verifying this as a disbursement message
	If CSES_DISB_check = "DISB" then call launch_selected_script(script_repository & "dail\cses-scrubber.vbs") 'If it's a disbursement message...
End if

'Disability certification ends in 60 days (loads DISA MESSAGE)
EMReadScreen DISA_check, 58, 6, 20
If DISA_check = "DISABILITY IS ENDING IN 60 DAYS - REVIEW DISABILITY STATUS" then call launch_selected_script(script_repository & "dail\disa-message.vbs")

'EMPS - ES Referral missing
EMReadScreen EMPS_ES_check, 52, 6, 20
If EMPS_ES_check = "EMPS:ES REFERRAL DATE IS BLANK FOR NON-EXEMPT PERSON" then call launch_selected_script(script_repository & "dail\es-referral-missing.vbs")

'EMPS - Financial Orientation date needed
EMReadScreen EMPS_Fin_Ori_check, 57, 6, 20
If EMPS_Fin_Ori_check = "REVIEW EMPS PANEL FOR FINANCIAL ORIENT DATE OR GOOD CAUSE" then call launch_selected_script(script_repository & "dail\financial-orientation-missing.vbs")

'Client can receive an FMED deduction for SNAP (loads FMED DEDUCTION)
EMReadScreen FMED_check, 59, 6, 20
If FMED_check = "MEMBER HAS TURNED 60 - NOTIFY ABOUT POSSIBLE FMED DEDUCTION" then call launch_selected_script(script_repository & "dail\fmed-deduction.vbs")

'Remedial care messages. May only happen at COLA (loads LTC - REMEDIAL CARE)
EMReadScreen remedial_care_check, 41, 6, 20
If remedial_care_check = "REF 01 PERSON HAS REMEDIAL CARE DEDUCTION" then call launch_selected_script(script_repository & "dail\ltc-remedial-care.vbs")

'New HIRE messages, client started a new job (loads NEW HIRE)
EMReadScreen HIRE_check, 15, 6, 20
If HIRE_check = "NEW JOB DETAILS" then call launch_selected_script(script_repository & "dail\new-hire.vbs")

'New HIRE messages, client started a new job (loads NEW HIRE)
EMReadScreen HIRE_check, 11, 6, 27
If HIRE_check = "JOB DETAILS" then call launch_selected_script(script_repository & "dail\new-hire-ndnh.vbs")

'Sends NOMI is DAIL generated by the REVS scrubber (loads SEND NOMI)
EMReadScreen NOMI_check, 11, 6, 20
If NOMI_check = "~*~*~CLIENT" then call launch_selected_script(script_repository &  "dail\send-nomi.vbs")

'SSI info received by agency (loads SDX INFO HAS BEEN STORED)
EMReadScreen SDX_check, 44, 6, 30
If SDX_check = "SDX INFORMATION HAS BEEN STORED - CHECK INFC" then 

	EMConnect ""
	EMSendKey "i" + "<enter>"
	
	EMWaitReady 0, 0
	EMSetCursor 20, 71
	EMSendKey "sdxs" + "<enter>"
	
	EMWaitReady 0, 0
	
	script_end_procedure("")

	' call launch_selected_script(script_repository & "dail\sdx-info-has-been-stored.vbs")
END IF

'Student income is ending (loads STUDENT INCOME)
EMReadScreen SCHL_check, 58, 6, 20
If SCHL_check = "STUDENT INCOME HAS ENDED - REVIEW FS AND/OR HC RESULTS/APP" then call launch_selected_script(script_repository & "dail\student-income.vbs")

'SSA info received by agency (loads TPQY RESPONSE)
EMReadScreen TPQY_check, 31, 6, 30
If TPQY_check = "TPQY RESPONSE RECEIVED FROM SSA" then call launch_selected_script(script_repository & "dail\tpqy-response.vbs")

'TYMA scrubber for agencies TIKLING TYMA as you go (loads TYMA Scrubber)
EMReadScreen TYMA_check, 23, 6, 20
IF TYMA_check = "~*~CONSIDER SENDING 1ST" THEN call launch_selected_script(script_repository & "dail\tyma-scrubber.vbs")
IF TYMA_check = "~*~CONSIDER SENDING 2ND" THEN Call launch_selected_script(script_repository & "dail\tyma-scrubber.vbs")
IF TYMA_check = "~*~2ND QUARTERLY REPORT" THEN call launch_selected_script(script_repository & "dail\tyma-scrubber.vbs")
IF TYMA_check = "~*~CONSIDER SENDING 3RD" THEN call launch_selected_script(script_repository & "dail\tyma-scrubber.vbs")
IF TYMA_check = "~*~3RD QUARTERLY REPORT" THEN call launch_selected_script(script_repository & "dail\tyma-scrubber.vbs")

'FS Eligibility Ending for ABAWD
EMReadScreen ABAWD_elig_end, 32, 6, 20
IF ABAWD_elig_end = "FS ABAWD ELIGIBILITY HAS EXPIRED" THEN CALL launch_selected_script(script_repository & "dail\abawd-fset-exemption-check.vbs")

'WAGE MATCH Scrubber
EMReadScreen wage_match, 4, 6, 6
IF wage_match = "WAGE" THEN CALL launch_selected_script(script_repository & "dail\wage-match-scrubber.vbs")

' EMReadScreen 
' if review_for_op = "REVIEW FOR POSSIBLE OVERPAYMENT" THEN CALL launch_selected_script(script_repository & "dail\possible_overpayment.vbs")

'NOW IF NO SCRIPT HAS BEEN WRITTEN FOR IT, THE DAIL SCRUBBER STOPS AND GENERATES A MESSAGE TO THE WORKER.----------------------------------------------------------------------------------------------------
MsgBox("You are not on a supported DAIL message. The script will now stop. " & vbNewLine & vbNewLine & "The message reads: " & full_message)
stopscript	