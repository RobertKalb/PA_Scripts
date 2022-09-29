'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - PARIS MATCH.vbs"
start_time = timer
STATS_counter = 1              'sets the stats counter at one
STATS_manualtime = 90          'manual run time in seconds
STATS_denomination = "C"      'C is for each case
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

'DIALOGS-------------------------------------------------------------
BeginDialog Paris_dialog, 0, 0, 231, 145, "Paris Dialog"
  EditBox 60, 5, 55, 15, Maxis_Case_number
  EditBox 170, 5, 25, 15, month_month
  EditBox 200, 5, 25, 15, year_year
  EditBox 60, 25, 55, 15, hhld_member_number
  EditBox 165, 25, 60, 15, state_state
  EditBox 50, 45, 65, 15, Programs_programs
  DropListBox 165, 45, 60, 15, "Select One..."+chr(9)+"UR"+chr(9)+"RV"+chr(9)+"FR"+chr(9)+"PR"+chr(9)+"HM"+chr(9)+"PC", code_used_dropdown
  OptionGroup RadioGroup1
    RadioButton 5, 70, 65, 10, "Match Resolved", match_resolved_radio
    RadioButton 95, 70, 85, 10, "Notice sent to client", notice_sent_radio
  EditBox 50, 85, 175, 15, other_notes
  EditBox 105, 105, 120, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 120, 125, 50, 15
    CancelButton 175, 125, 50, 15
  Text 5, 10, 45, 10, "Case number:"
  Text 125, 10, 40, 10, "Month/Year:"
  Text 5, 30, 50, 10, "HHLD Member:"
  Text 125, 30, 25, 10, "State:"
  Text 10, 50, 35, 10, "Programs:"
  Text 125, 50, 40, 10, "Code Used:"
  Text 75, 70, 15, 10, "-or-"
  Text 5, 90, 40, 10, "Other notes: "
  Text 40, 110, 60, 10, "Worker Signature:"
EndDialog

'--THE SCRIPT----------------------------------------------------
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)

Do
	DO
		Err_msg = ""
		Dialog Paris_dialog
		cancel_confirmation
			If Maxis_Case_number = "" or IsNumeric(case_number) = False or len(case_number) > 8 THEN err_msg = err_msg & vbNewLine & "*Please enter a valid case number"
			If month_month = "" THEN err_msg = err_msg & vbNewLine & "*Please enter the month of the Paris Match"
			If year_year = "" THEN err_msg = err_msg & vbNewLine & "*Please enter the year of the Paris Match"
			If hhld_member_number = "" THEN err_msg = err_msg & vbNewLine & "*Please enter the household member"
			If state_state = "" THEN err_msg = err_msg & vbNewLine & "*Please enter the state"
			If programs_programs = "" THEN err_msg = err_msg & vbNewLine & "*Please enter the program"
			If code_used_dropdown = "Select One..." THEN err_msg = err_msg & vbNewLine & "*Please select the code used"
			If worker_signature = "" THEN err_msg = err_msg & vbNewLine & "*Please sign your case note"
			If err_msg <> "" Then msgbox "***NOTICE!!!***" & vbNewLine & err_msg & vbNewLine & vbNewLine & "Please resolve for the script to continue"
	Loop until err_msg = ""
	CALL check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = false

Dim Paris_match_header
If match_resolved_radio = checked THEN Paris_match_header = Paris_match_header & "- Resolved"
If notice_sent_radio = checked THEN Paris_match_header = Paris_match_header & "- Notice sent to client"

call check_for_MAXIS(True)	'checking for an active MAXIS session

'Writing the case note to MAXIS---
call start_a_blank_CASE_NOTE
call write_variable_in_case_note("PARIS Match" & paris_match_header)
call write_bullet_and_variable_in_case_note("Household Member", hhld_member_number)
call write_bullet_and_variable_in_case_note("Month/Year", month_month & "/" & year_year)
call write_bullet_and_variable_in_case_note("State", state_state)
call write_bullet_and_variable_in_case_note("Programs", programs_programs)
call write_bullet_and_variable_in_case_note("Code Used", code_used_dropdown)
call write_bullet_and_variable_in_case_note("Notes", other_notes)
CALL write_variable_in_CASE_NOTE ("---")
call write_variable_in_case_note(worker_signature)

script_end_procedure("")
