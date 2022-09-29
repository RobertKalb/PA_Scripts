'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - FOSTER CARE REVIEW.vbs"
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
BeginDialog FC_HC_review_dialog, 0, 0, 256, 250, "FOSTER CARE HC REVIEW"
  EditBox 65, 5, 65, 15, MAXIS_case_number
  EditBox 65, 25, 65, 15, Received
  EditBox 65, 45, 65, 15, Completed_By
  EditBox 130, 70, 105, 15, Social_Worker_or_Probation_Officer
  EditBox 105, 90, 85, 15, Extended_Foster_Care_Date
  EditBox 40, 125, 55, 15, Income
  EditBox 40, 145, 70, 15, Results
  EditBox 75, 190, 110, 15, Worker_Signature
  ButtonGroup ButtonPressed
    OkButton 125, 230, 50, 15
    CancelButton 190, 230, 50, 15
  Text 5, 5, 55, 10, "Case number:"
  Text 5, 25, 45, 10, "Received: "
  Text 5, 45, 60, 10, "Completed By: "
  Text 5, 70, 120, 10, "Social Worker or Probation Officer:"
  Text 5, 95, 95, 15, "Extended Foster Care Date: "
  Text 5, 125, 35, 15, "Income: "
  Text 5, 150, 30, 15, "Results:"
  Text 5, 190, 65, 10, "Worker Signature: "
EndDialog



'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------
'connecting to BlueZone, and grabbing the case number
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)

'calling the dialog---------------------------------------------------------------------------------------------------------------
DO
	Dialog FC_HC_review_dialog
	IF buttonpressed = 0 THEN stopscript
	IF MAXIS_case_number = "" THEN MsgBox "You must have a case number to continue!"
	IF worker_signature = "" THEN MsgBox "You must enter a worker signature."
LOOP until MAXIS_case_number <> "" and worker_signature <> ""

'checking for an active MAXIS session
CALL check_for_MAXIS(FALSE)

'The case note---------------------------------------------------------------------------------------------------------------------
start_a_blank_CASE_NOTE
CALL write_variable_in_CASE_NOTE("***Foster Care HC REVIEW***")
CALL write_bullet_and_variable_in_CASE_NOTE("Received", Received)
CALL write_bullet_and_variable_in_CASE_NOTE("Completed By", Completed_By)
CALL write_bullet_and_variable_in_CASE_NOTE("Social Worker or Probation Officer", Social_Worker_or_Probation_Officer)
CALL write_bullet_and_variable_in_CASE_NOTE("Extended Foster Care Date", Extended_Foster_Care_Date)
CALL write_bullet_and_variable_in_CASE_NOTE("Income", Income)
CALL write_bullet_and_variable_in_CASE_NOTE("Results", Results)
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE(worker_signature)
Script_end_procedure("")
