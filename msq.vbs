'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - MSQ.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 300          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block==========================================================================================================

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

'THE DIALOG----------------------------------------------------------------------------------------------------------
BeginDialog msq_dialog, 0, 0, 321, 125, "MSQ"
  EditBox 80, 5, 70, 15, MAXIS_case_number
  EditBox 75, 30, 70, 15, member_injured
  EditBox 205, 30, 70, 15, injury_date
  EditBox 75, 65, 175, 15, other_notes
  EditBox 80, 95, 80, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 200, 95, 50, 15
    CancelButton 255, 95, 50, 15
  Text 5, 70, 70, 10, "Action Taken/Notes:"
  Text 165, 35, 40, 10, "Injury Date:"
  Text 5, 35, 70, 10, "HH Member Injured:"
  Text 5, 100, 70, 10, "Sign your Case Note:"
  Text 5, 10, 70, 10, "Maxis Case Number:"
  Text 75, 45, 40, 10, "(Ex: 01, 02)"
  Text 205, 45, 70, 10, "(Ex: MM/DD/YY)"
EndDialog


'THE SCRIPT--------------------------------------------------------------------------------------------------------------

'Connects to BLUEZONE
EMConnect ""

'Grabs the MAXIS case number
CALL MAXIS_case_number_finder(MAXIS_case_number)

'Shows dialog
DO
	err_msg = ""
	Dialog msq_dialog
		IF ButtonPressed = 0 THEN StopScript
		IF IsNumeric(MAXIS_case_number) = FALSE THEN err_msg = err_msg & vbCr & "* You must type a valid numeric case number."
		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* You must sign your case note!"
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
LOOP UNTIL err_msg = ""

'Checks Maxis for password prompt
CALL check_for_MAXIS(True)

'The script reads what member number was manually entered, and navigates to that member's STAT/ACCI panel
CALL navigate_to_MAXIS_screen("STAT", "ACCI")
EMWriteScreen member_injured, 20, 76
EMWriteScreen "nn", 20, 79
transmit

EMWriteScreen "n", 8, 75

'Writes 13 in Accident Type field
EMWriteScreen "13", 6, 47

'Writes the Injury Date in the Injury date field
CALL create_MAXIS_friendly_date(injury_date, 0, 6, 73)

'Writes N in the Med Cooperation field
EMWriteScreen "N", 7, 47

'Writes N in the Good cause field
EMWriteScreen "N", 7, 73

'Writes a N in Pend Litigation
EMWritescreen "N", 9, 47

'Opens new case note
start_a_blank_case_note


'Writes the Case Note
CALL write_variable_in_case_note("*** MSQ Form ***")
CALL write_bullet_and_variable_in_case_note("Household Member Injured", member_injured)
CALL write_bullet_and_variable_in_case_note("Injury Date", injury_date)
CALL write_bullet_and_variable_in_CASE_NOTE("Actions Taken/Notes", other_notes)
CALL write_variable_in_case_note("---")
CALL write_variable_in_case_note(worker_signature)

script_end_procedure("Success! Remember to update MMIS.")
