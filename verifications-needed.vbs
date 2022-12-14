'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - VERIFICATIONS NEEDED.vbs"
start_time = timer
STATS_counter = 1         'sets the stats counter to 1
STATS_manualtime = 210    'sets the manual run time
STATS_denomination = "C"  'C is for case
'END OF stats block=========================================================================================================

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

' 'CHANGELOG BLOCK ===========================================================================================================
' 'Starts by defining a changelog array
' changelog = array()
' 
' 'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
' 'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
' call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")
' 
' 'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
' changelog_display
' 'END CHANGELOG BLOCK =======================================================================================================

'THIS SCRIPT IS BEING USED IN A WORKFLOW SO DIALOGS ARE NOT NAMED
'DIALOGS MAY NOT BE DEFINED AT THE BEGINNING OF THE SCRIPT BUT WITHIN THE SCRIPT FILE

'THE SCRIPT--------------------------------------------------------------------------------------------------
'Asks if this is a LTC case or not. LTC has a different dialog. The if...then logic will be put in the do...loop.
LTC_case = MsgBox("Is this a Long Term Care case? LTC cases have a few more options on their dialog.", vbYesNoCancel)
If LTC_case = vbCancel then stopscript

'Connects to BlueZone
EMConnect ""
'Calls a MAXIS case number
call MAXIS_case_number_finder(MAXIS_case_number)

'Shows dialog. Requires a case number, checks for an active MAXIS session, and checks that it can add/update a case note before proceeding.
If LTC_case = vbYes then 									'Shows dialog if LTC
	DO
		Do
			Do
				'Dialog for LTC cases is defined here - not named
				BeginDialog , 0, 0, 351, 435, "Verifs needed (LTC) dialog"
				  EditBox 55, 5, 70, 15, MAXIS_case_number
				  EditBox 250, 5, 60, 15, verif_due_date
				  ButtonGroup ButtonPressed
					PushButton 315, 10, 30, 10, "CD+10", CD_plus_10_button
				  EditBox 30, 40, 315, 15, FACI
				  EditBox 30, 60, 130, 15, JOBS
				  EditBox 205, 60, 140, 15, BUSI_RBIC
				  EditBox 45, 80, 300, 15, UNEA_01
				  EditBox 75, 100, 270, 15, UNEA_other_membs
				  EditBox 45, 120, 300, 15, ACCT_01
				  EditBox 75, 140, 270, 15, ACCT_other_membs
				  EditBox 45, 160, 300, 15, SECU_01
				  EditBox 75, 180, 270, 15, SECU_other_membs
				  EditBox 30, 200, 315, 15, CARS
				  EditBox 30, 220, 315, 15, REST
				  EditBox 50, 240, 295, 15, OTHR
				  EditBox 30, 260, 315, 15, SHEL
				  EditBox 30, 280, 315, 15, INSA
				  EditBox 70, 300, 275, 15, medical_expenses
				  EditBox 50, 320, 295, 15, veterans_info
				  EditBox 50, 340, 295, 15, other_proofs
				  CheckBox 5, 360, 240, 10, "Check here if you sent form DHS-2919A (Verification Request Form - A).", verif_A_check
				  CheckBox 5, 375, 240, 10, "Check here if you sent form DHS-2919B (Verification Request Form - B).", verif_B_check
				  CheckBox 5, 390, 165, 10, "Sent form to AREP?", sent_arep_checkbox
				  CheckBox 5, 405, 95, 10, "Signature page needed?", signature_page_needed_check
				  CheckBox 5, 420, 130, 10, "Check here to TIKL out for this case.", TIKL_check
				  EditBox 285, 385, 60, 15, worker_signature
				  ButtonGroup ButtonPressed
					OkButton 240, 405, 50, 15
					CancelButton 295, 405, 50, 15
				  Text 150, 10, 100, 10, "Verifs due by (MM/DD/YYYY):"
				  Text 5, 25, 300, 10, "If you aren't requesting something, leave that section blank. That way it doesn't case note."
				  Text 5, 45, 25, 10, "FACI:"
				  Text 5, 65, 25, 10, "JOBS:"
				  Text 165, 65, 40, 10, "BUSI/RBIC:"
				  Text 5, 85, 35, 10, "UNEA 01:"
				  Text 5, 105, 65, 10, "UNEA other membs:"
				  Text 5, 125, 35, 10, "ACCT 01:"
				  Text 5, 145, 65, 10, "ACCT other membs:"
				  Text 5, 165, 35, 10, "SECU 01:"
				  Text 5, 185, 70, 10, "SECU other membs:"
				  Text 5, 205, 25, 10, "CARS:"
				  Text 5, 225, 25, 10, "REST:"
				  Text 5, 245, 45, 10, "Burial/OTHR:"
				  Text 5, 265, 25, 10, "SHEL:"
				  Text 5, 285, 25, 10, "INSA:"
				  Text 5, 305, 65, 10, "Medical expenses:"
				  Text 5, 345, 45, 10, "Other proofs:"
				  Text 220, 390, 60, 10, "worker signature:"
				  Text 5, 10, 50, 10, "Case number:"
				  Text 5, 325, 45, 10, "Veteran info:"
				EndDialog
				DIALOG 					'Calling a dialog without a assigned variable will call the most recently defined dialog
				cancel_confirmation													'quits if cancel is pressed
				If buttonpressed = CD_plus_10_button then verif_due_date = dateadd("d", 10, date) & ""		'Fills in current date + 10 if you press the button.
			Loop until buttonpressed = OK																	'Loops until you press OK
			If MAXIS_case_number = "" then MsgBox "You must have a case number to continue!"		'Yells at you if you don't have a case number
		Loop until MAXIS_case_number <> ""														'Loops until that case number exists
		call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
	LOOP UNTIL are_we_passworded_out = false														'Loops until that case number exists
ELSEIF LTC_case = vbNo then							'Shows dialog if not LTC
	DO
		Do
			Do
				'Dialog for all other cases is defined here
				BeginDialog , 0, 0, 351, 360, "Verifs needed"
				  EditBox 55, 5, 70, 15, MAXIS_case_number
				  EditBox 250, 5, 60, 15, verif_due_date
				  EditBox 30, 35, 315, 15, ADDR
				  EditBox 70, 55, 275, 15, SCHL
				  EditBox 30, 75, 315, 15, DISA
				  EditBox 30, 95, 315, 15, JOBS
				  EditBox 30, 115, 315, 15, BUSI
				  EditBox 30, 135, 315, 15, UNEA
				  EditBox 30, 155, 315, 15, ACCT
				  EditBox 55, 175, 290, 15, other_assets
				  EditBox 30, 195, 315, 15, SHEL
				  EditBox 30, 215, 315, 15, INSA
				  EditBox 50, 235, 295, 15, other_proofs
				  CheckBox 5, 260, 240, 10, "Check here if you sent form DHS-2919A (Verification Request Form - A).", verif_A_check
				  CheckBox 5, 275, 240, 10, "Check here if you sent form DHS-2919B (Verification Request Form - B).", verif_B_check
				  CheckBox 5, 290, 240, 15, "Check here if these are postponed verifications for expedited SNAP.  ", postponed_check
				  CheckBox 5, 310, 175, 10, "Sent form to AREP?", sent_arep_checkbox
				  CheckBox 5, 325, 95, 10, "Signature page needed?", signature_page_needed_check
				  CheckBox 5, 340, 130, 10, "Check here to TIKL out for this case.", TIKL_check
				  EditBox 285, 315, 60, 15, worker_signature
				  ButtonGroup ButtonPressed
					OkButton 240, 340, 50, 15
					CancelButton 295, 340, 50, 15
					PushButton 315, 10, 30, 10, "CD+10", CD_plus_10_button
				  Text 5, 10, 50, 10, "Case number:"
				  Text 150, 10, 100, 10, "Verifs due by (MM/DD/YYYY):"
				  Text 5, 25, 300, 10, "If you aren't requesting something, leave that section blank. That way it doesn't case note."
				  Text 5, 40, 25, 10, "ADDR:"
				  Text 5, 60, 60, 10, "SCHL/STIN/STEC:"
				  Text 5, 80, 25, 10, "DISA:"
				  Text 5, 100, 25, 10, "JOBS:"
				  Text 5, 120, 20, 10, "BUSI:"
				  Text 5, 140, 25, 10, "UNEA:"
				  Text 5, 160, 25, 10, "ACCT:"
				  Text 5, 180, 50, 10, "Other assets:"
				  Text 5, 200, 25, 10, "SHEL:"
				  Text 5, 220, 25, 10, "INSA:"
				  Text 5, 240, 45, 10, "Other proofs:"
				  Text 215, 320, 70, 10, "Sign your case note:"
				 EndDialog
				DIALOG 					'Calling a dialog without a assigned variable will call the most recently defined dialog
				cancel_confirmation													'quits if cancel is pressed
				If buttonpressed = CD_plus_10_button then verif_due_date = dateadd("d", 10, date) & ""		'Fills in current date + 10 if you press the button.
			Loop until buttonpressed = OK																	'Loops until you press OK
			If MAXIS_case_number = "" then MsgBox "You must have a case number to continue!"		'Yells at you if you don't have a case number
		Loop until MAXIS_case_number <> ""														'Loops until that case number exists
		call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
	LOOP UNTIL are_we_passworded_out = false
END IF

'checking for an active MAXIS session
Call check_for_MAXIS(False)


'THE CASE NOTE----------------------------------------------------------------------------------------------------
'Writes a new line, then writes each additional line if there's data in the dialog's edit box (uses if/then statement to decide).
start_a_blank_CASE_NOTE
IF postponed_check = checked THEN
	call write_variable_in_case_note(">>>POSTPONED VERIFICATIONS REQUESTED FOR EXP SNAP<<<")
ELSE
	call write_variable_in_case_note(">>>Verifications Requested<<<")
END IF
call write_bullet_and_variable_in_case_note("Verif due date", verif_due_date)
call write_bullet_and_variable_in_case_note("ADDR", ADDR)
call write_bullet_and_variable_in_case_note("FACI", FACI)
call write_bullet_and_variable_in_case_note("SCHL/STIN/STEC", SCHL)
call write_bullet_and_variable_in_case_note("DISA", DISA)
call write_bullet_and_variable_in_case_note("JOBS", JOBS)
call write_bullet_and_variable_in_case_note("BUSI", BUSI)
call write_bullet_and_variable_in_case_note("BUSI/RBIC", BUSI_RBIC)
call write_bullet_and_variable_in_case_note("UNEA", UNEA)
call write_bullet_and_variable_in_case_note("UNEA (MEMB 01)", UNEA_01)
call write_bullet_and_variable_in_case_note("UNEA (other membs)", UNEA_other_membs)
call write_bullet_and_variable_in_case_note("ACCT", ACCT)
call write_bullet_and_variable_in_case_note("ACCT (MEMB 01)", ACCT_01)
call write_bullet_and_variable_in_case_note("ACCT (other membs)", ACCT_other_membs)
call write_bullet_and_variable_in_case_note("SECU (MEMB 01)", SECU_01)
call write_bullet_and_variable_in_case_note("SECU (other membs)", SECU_other_membs)
call write_bullet_and_variable_in_case_note("CARS", CARS)
call write_bullet_and_variable_in_case_note("REST", REST)
call write_bullet_and_variable_in_case_note("Burial/OTHR", OTHR)
call write_bullet_and_variable_in_case_note("Other assets", other_assets)
call write_bullet_and_variable_in_case_note("SHEL", SHEL)
call write_bullet_and_variable_in_case_note("INSA", INSA)
call write_bullet_and_variable_in_case_note("Veteran's info", veterans_info)
call write_bullet_and_variable_in_case_note("Medical expenses", medical_expenses)
call write_bullet_and_variable_in_case_note("Other proofs", other_proofs)
IF verif_A_check = checked THEN write_variable_in_CASE_NOTE("* Verification request form A sent.")
IF verif_B_check = checked THEN write_variable_in_CASE_NOTE("* Verification request form B sent.")
IF sent_arep_checkbox = checked THEN write_variable_in_CASE_NOTE("* Forms sent to AREP.")
IF signature_page_needed_check = checked THEN write_variable_in_CASE_NOTE("* Signature page needed.")
Call write_variable_in_case_note("---")
call write_variable_in_CASE_NOTE(worker_signature)

'THE TIKL----------------------------------------------------------------------------------------------------
'If TIKL_check isn't checked this is the end
If TIKL_check = unchecked then script_end_procedure("")

'Navigating to DAIL/WRIT
call navigate_to_MAXIS_screen("dail", "writ")

'If the date in Verif due date is a date, it'll fill that date in on the TIKL.
If IsDate(verif_due_date) = True then call create_MAXIS_friendly_date(verif_due_date, 0, 5, 18)

'Script ends
script_end_procedure("Success! Case note made. You may TIKL when ready. If you filled in a verif due date, it should be autofilled in this TIKL.")
