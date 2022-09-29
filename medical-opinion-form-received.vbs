'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - MEDICAL OPINION FORM RECEIVED.vbs"
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
BeginDialog MOF_recd, 0, 0, 186, 265, "Medical Opinion Form Received"
  EditBox 55, 5, 100, 15, MAXIS_case_number
  EditBox 55, 25, 95, 15, date_recd
  EditBox 80, 45, 90, 15, HH_Member
  CheckBox 20, 65, 85, 10, "Client signed release?", client_release
  EditBox 75, 80, 100, 15, last_exam_date
  EditBox 90, 100, 85, 15, doctor_date
  EditBox 70, 120, 105, 15, condition_will_last
  EditBox 85, 160, 90, 15, ability_to_work
  EditBox 50, 180, 125, 15, other_notes
  EditBox 50, 200, 125, 15, action_taken
  EditBox 70, 220, 105, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 70, 240, 50, 15
    CancelButton 125, 240, 50, 15
  Text 5, 10, 50, 10, "Case Number: "
  Text 5, 30, 50, 10, "Date received: "
  Text 5, 50, 70, 10, "HHLD Member name"
  Text 5, 85, 65, 10, "Date of last exam: "
  Text 5, 105, 80, 10, "Date doctor signed form: "
  Text 5, 125, 65, 10, "Condition will last:"
  Text 5, 145, 175, 10, "Do not enter diagnosis in case notes per PQ #16506."
  Text 5, 165, 75, 10, "Client's ability to work: "
  Text 5, 185, 40, 10, "Other notes: "
  Text 5, 205, 45, 10, "Action taken: "
  Text 5, 225, 60, 10, "Worker Signature: "
EndDialog




'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------
'connecting to BlueZone, and grabbing the case number
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)


'calling the dialog---------------------------------------------------------------------------------------------------------------
DO
	Err_msg = ""
	Dialog MOF_recd
	IF buttonpressed = 0 THEN stopscript
	IF MAXIS_case_number = "" THEN err_msg = err_msg & vbNewLine & "*You must enter a case number"
	IF worker_signature = "" THEN err_msg = err_msg & vbNewLine & "You must enter a worker signature."
	If HH_Member = "" Then err_msg = err_msg & vbNewLine & "*You must enter the household member"
	If err_msg <> "" Then msgbox "***NOTICE!!!***" & vbNewLine & err_msg & vbNewLine & vbNewLine & "Please resolve for the script to continue"
LOOP until err_msg = ""


'checking for an active MAXIS session
CALL check_for_MAXIS(FALSE)

CALL navigate_to_MAXIS_screen("STAT", "PROG")  'checking for stat to remind worker about WREG/ABAWD
EMReadScreen SNAP_ACTV, 4, 10, 74


'The case note---------------------------------------------------------------------------------------------------------------------
start_a_blank_CASE_NOTE
CALL write_variable_in_CASE_NOTE("***Medical Opinion Form Rec'd " & date_recd & "***")
Call write_bullet_and_variable_in_CASE_NOTE("Household Member", HH_Member)
IF client_release = checked THEN CALL write_variable_in_CASE_NOTE ("* Client signed release on MOF.")
CALL write_bullet_and_variable_in_CASE_NOTE("Date of last examination", last_exam_date)
CALL write_bullet_and_variable_in_CASE_NOTE("Doctor signed form", doctor_date)
CALL write_bullet_and_variable_in_CASE_NOTE("Condition will last", condition_will_last)
CALL write_bullet_and_variable_in_CASE_NOTE("Ability to work", ability_to_work)
CALL write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
CALL write_bullet_and_variable_in_CASE_NOTE("Action taken", action_taken)
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE(worker_signature)

IF SNAP_ACTV = "ACTV" or SNAP_ACTV = "PEND" THEN MSGBOX "Please remember to update WREG and client's ABAWD status accordingly."  'Adds message box to remind worker to update WREG and ABAWD if SNAP is ACTV or pending

Script_end_procedure("")
