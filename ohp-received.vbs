'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - OHP RECEIVED.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 90           'manual run time in seconds
STATS_denomination = "C"        'C is for each case
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

'Dialog---------------------------------------------------------------------------------------------------------------------------
BeginDialog OHP_dialog, 0, 0, 301, 160, "OHP received"
  EditBox 90, 5, 75, 15, MAXIS_case_number
  EditBox 145, 25, 65, 15, OOHP_date
  EditBox 65, 45, 90, 15, Date_change
  EditBox 65, 70, 145, 15, Change
  EditBox 65, 90, 150, 15, Action_taken
  EditBox 80, 115, 110, 15, Worker_Signature
  ButtonGroup ButtonPressed
    OkButton 160, 140, 50, 15
    CancelButton 230, 140, 50, 15
  Text 5, 5, 70, 10, "Case number:"
  Text 5, 25, 130, 10, "Out of home placement form received:"
  Text 5, 45, 60, 10, "Date of change:"
  Text 5, 70, 45, 10, "Change: "
  Text 5, 95, 50, 15, "Action Taken: "
  Text 5, 115, 65, 10, "Worker Signature: "
EndDialog



'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------
'connecting to BlueZone, and grabbing the case number
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)

'calling the dialog---------------------------------------------------------------------------------------------------------------
DO
	Dialog OHP_dialog
	IF buttonpressed = 0 THEN stopscript
	IF MAXIS_case_number = "" THEN MsgBox "You must have a case number to continue!"
	IF Worker_Signature = "" THEN MsgBox "You must enter a worker signature."
LOOP until MAXIS_case_number <> "" and Worker_Signature <> ""

'checking for an active MAXIS session
CALL check_for_MAXIS(FALSE)

'The case note---------------------------------------------------------------------------------------------------------------------
start_a_blank_CASE_NOTE
CALL write_variable_in_CASE_NOTE("***Out Of Home Placement Received***")
CALL write_bullet_and_variable_in_CASE_NOTE("OHP date Received", OOHP_date)
CALL write_bullet_and_variable_in_CASE_NOTE("Date of change", Date_change)
CALL write_bullet_and_variable_in_CASE_NOTE("Change", Change)
CALL write_bullet_and_variable_in_CASE_NOTE("Action taken", Action_taken)
CALL write_variable_in_CASE_NOTE(Worker_Signature)

Script_end_procedure("")
