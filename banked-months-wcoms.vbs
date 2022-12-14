'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "NOTICES - BANKED MONTHS WCOMS.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 90                               'manual run time in seconds
STATS_denomination = "C"       'C is for each CASE
'END OF stats block==============================================================================================

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
' 'CHANGELOG BLOCK ===========================================================================================================
' 'Starts by defining a changelog array
' changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
'call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
'changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'Dialogs
BeginDialog case_number_dlg, 0, 0, 211, 80, "Case Number Dialog"
  EditBox 70, 10, 60, 15, MAXIS_case_number
  EditBox 70, 30, 30, 15, approval_month
  EditBox 160, 30, 30, 15, approval_year
  ButtonGroup ButtonPressed
    OkButton 45, 55, 50, 15
    CancelButton 105, 55, 50, 15
  Text 10, 15, 55, 10, "Case Number: "
  Text 10, 35, 55, 10, "Approval Month:"
  Text 105, 35, 50, 10, "Approval Year:"
EndDialog


BeginDialog banked_months_menu_dialog, 0, 0, 356, 140, "Banked Months WCOMs"
  ButtonGroup ButtonPressed
    PushButton 10, 25, 90, 10, "All Banked Months Used", banked_months_used_button
    PushButton 10, 50, 90, 10, "Banked Months Notifier", banked_months_notifier
    PushButton 10, 75, 90, 10, "Closing for E/T Non-Coop", e_t_non_coop_button
    CancelButton 300, 120, 50, 15
  Text 110, 25, 230, 20, "-- Use this script when a client's SNAP is closing because they used all their eligible banked months."
  Text 110, 50, 230, 20, "-- Use this script to add a WCOM to a notice notifying the client they may be eligible for banked months."
  Text 110, 75, 235, 25, "-- Use this script to add a WCOM to a client's closing notice to inform them they are closing on banked months for Employment Services Non-Coop."
  GroupBox 5, 10, 345, 90, "WCOM"
EndDialog





'--- The script -----------------------------------------------------------------------------------------------------------------

EMConnect ""


call MAXIS_case_number_finder(MAXIS_case_number)
approval_month = DatePart("M", (DateAdd("M", 1, date)))
IF len(approval_month) = 1 THEN
	approval_month = "0" & approval_month
ELSE
	approval_month = Cstr(approval_month)
END IF
approval_year = Right(DatePart("YYYY", (DateAdd("M", 1, date))), 2)

DO
	err_msg = ""
	dialog case_number_dlg
	cancel_confirmation
	IF MAXIS_case_number = "" THEN err_msg = "* Please enter a case number" & vbNewLine
	IF len(approval_month) <> 2 THEN err_msg = err_msg & "* Please enter your month in MM format." & vbNewLine
	IF len(approval_year) <> 2 THEN err_msg = err_msg & "* Please enter your year in YY format." & vbNewLine
	IF err_msg <> "" THEN msgbox err_msg
LOOP until err_msg = ""

CALL check_for_MAXIS(false)

DIALOG banked_months_menu_dialog
	cancel_confirmation

	'This is the WCOM for when the client has used all their banked months.
	IF ButtonPressed = banked_months_used_button THEN
		call navigate_to_MAXIS_screen("spec", "wcom")

		EMWriteScreen approval_month, 3, 46
		EMWriteScreen approval_year, 3, 51
		transmit

		DO 								'This DO/LOOP resets to the first page of notices in SPEC/WCOM
			EMReadScreen more_pages, 8, 18, 72
			IF more_pages = "MORE:  -" THEN PF7
		LOOP until more_pages <> "MORE:  -"

		read_row = 7
		DO
			waiting_check = ""
			EMReadscreen prog_type, 2, read_row, 26
			EMReadscreen waiting_check, 7, read_row, 71 'finds if notice has been printed
			If waiting_check = "Waiting" and prog_type = "FS" THEN 'checking program type and if it's been printed
				EMSetcursor read_row, 13
				EMSendKey "x"
				Transmit
				pf9
				EMSetCursor 03, 15
				CALL write_variable_in_SPEC_MEMO("You have been receiving SNAP banked months. Your SNAP is closing for using all available banked months. If you meet one of the exemptions listed above AND all other eligibility factors you may still be eligible for SNAP. Please contact your financial worker if you have questions.")
				PF4
				PF3
				WCOM_count = WCOM_count + 1
				exit do
			ELSE
				read_row = read_row + 1
			END IF
			IF read_row = 18 THEN
				PF8          'Navigates to the next page of notices.  DO/LOOP until read_row = 18
				read_row = 7
			End if
		LOOP until prog_type = "  "

		wcom_type = "all banked months"

	'This is the WCOM for when the client is closing for ABAWD and is being notified that they could be eligible for banked months.
	ELSEIF ButtonPressed = banked_months_notifier THEN
		call navigate_to_MAXIS_screen("spec", "wcom")

		EMWriteScreen approval_month, 3, 46
		EMWriteScreen approval_year, 3, 51
		transmit

		DO 								'This DO/LOOP resets to the first page of notices in SPEC/WCOM
			EMReadScreen more_pages, 8, 18, 72
			IF more_pages = "MORE:  -" THEN PF7
		LOOP until more_pages <> "MORE:  -"

		read_row = 7
		DO
			waiting_check = ""
			EMReadscreen prog_type, 2, read_row, 26
			EMReadscreen waiting_check, 7, read_row, 71 'finds if notice has been printed
			If waiting_check = "Waiting" and prog_type = "FS" THEN 'checking program type and if it's been printed
				EMSetcursor read_row, 13
				EMSendKey "x"
				Transmit
				pf9
				EMSetCursor 03, 15
				CALL write_variable_in_SPEC_MEMO("You have used all of your available ABAWD months. You may be eligible for SNAP banked months if you are cooperating with Employment Services. Please contact your financial worker if you have questions.")
				PF4
				PF3
				WCOM_count = WCOM_count + 1
				exit do
			ELSE
				read_row = read_row + 1
			END IF
			IF read_row = 18 THEN
				PF8          'Navigates to the next page of notices.  DO/LOOP until read_row = 18
				read_row = 7
			End if
		LOOP until prog_type = "  "

		wcom_type = "banked months notifier"

	'This is the WCOM for when the client is closing on banked months for E&T Non-Coop
	ELSEIF ButtonPressed = e_t_non_coop_button THEN

		DO
			hh_member = InputBox("Please enter the name of the client that is closing for E&T Non-Coop...")
			confirmation_msg = MsgBox("Please confirm to add the client's name to the WCOM: " & vbCr & vbCr & hh_member & " is closing on banked months for SNAP E&T Non-Cooperation." & vbCr & vbCr & "Is this correct? Press YES to continue. Press NO to re-enter the client's name. Press CANCEL to stop the script.", vbYesNoCancel)
			IF confirmation_msg = vbCancel THEN stopscript
		LOOP UNTIL confirmation_msg = vbYes

		call navigate_to_MAXIS_screen("spec", "wcom")

		EMWriteScreen approval_month, 3, 46
		EMWriteScreen approval_year, 3, 51
		transmit

		DO 								'This DO/LOOP resets to the first page of notices in SPEC/WCOM
			EMReadScreen more_pages, 8, 18, 72
			IF more_pages = "MORE:  -" THEN PF7
		LOOP until more_pages <> "MORE:  -"

		read_row = 7
		DO
			waiting_check = ""
			EMReadscreen prog_type, 2, read_row, 26
			EMReadscreen waiting_check, 7, read_row, 71 'finds if notice has been printed
			If waiting_check = "Waiting" and prog_type = "FS" THEN 'checking program type and if it's been printed
				EMSetcursor read_row, 13
				EMSendKey "x"
				Transmit
				pf9
				EMSetCursor 03, 15
				CALL write_variable_in_SPEC_MEMO("You have been receiving SNAP banked months. Your SNAP case is closing because " & hh_member & " did not meet the requirements of working with Employment and Training. If you feel you have Good Cause for not cooperating with this requirement please contact your financial worker before your SNAP closes. If your SNAP closes for not cooperating with Employment and Training you will not be eligible for future banked months. If you meet an exemption listed above AND all other eligibility factors you may be eligible for SNAP. If you have questions please contact your financial worker.")
				PF4
				PF3
				WCOM_count = WCOM_count + 1
				exit do
			ELSE
				read_row = read_row + 1
			END IF
			IF read_row = 18 THEN
				PF8          'Navigates to the next page of notices.  DO/LOOP until read_row = 18
				read_row = 7
			End if
		LOOP until prog_type = "  "

		wcom_type = "non coop"
	END IF

'Outcome ---------------------------------------------------------------------------------------------------------------------

If WCOM_count = 0 THEN  'if no waiting FS notice is found
	script_end_procedure("No Waiting FS elig results were found in this month for this HH member.")
ELSE 					'If a waiting FS notice is found
	'Case note
	start_a_blank_case_note
	call write_variable_in_CASE_NOTE("---WCOM added regarding banked months---")
	IF wcom_type = "all banked months" THEN
		CALL write_variable_in_CASE_NOTE("* WCOM added because client all eligible banked months have been used.")
	ELSEIF wcom_type = "non coop" THEN
		CALL write_variable_in_CASE_NOTE("* Banked months ending for SNAP E & T non-coop.")
	ELSEIF wcom_type = "banked months notifier" THEN
		CALL write_variable_in_CASE_NOTE("* Client has used ABAWD counted months and MAY be eligible for banked months. Eligibility questions should be directed to financial worker.")
	END IF

	call write_variable_in_CASE_NOTE("---")
	IF worker_signature <> "" THEN
		call write_variable_in_CASE_NOTE(worker_signature)
	ELSE
		worker_signature = InputBox("Please sign your case note...")
		CALL write_variable_in_CASE_NOTE(worker_signature)
	END IF
END IF

script_end_procedure("")
