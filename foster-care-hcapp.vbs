'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - FOSTER CARE HCAPP.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 250         'manual run time in seconds
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
BeginDialog Foster_Care_HCAPP, 0, 0, 276, 270, "Foster Care HCAPP "
  EditBox 60, 5, 65, 15, MAXIS_case_number
  EditBox 190, 5, 75, 15, Completed_By
  EditBox 100, 25, 60, 15, Date_Received_By_Agency
  EditBox 105, 45, 55, 15, Date_of_Agency_Responsiblity
  EditBox 30, 65, 130, 15, IV_E
  EditBox 60, 85, 100, 15, SWKR_or_PO
  EditBox 35, 105, 125, 15, AREP
  EditBox 40, 125, 225, 15, Income
  EditBox 70, 145, 195, 15, Retro_Requested
  EditBox 45, 170, 65, 15, HHcomp
  EditBox 155, 170, 115, 15, Assets
  EditBox 35, 190, 235, 15, OHC
  EditBox 95, 210, 175, 15, Verifications_Requested
  EditBox 65, 230, 205, 15, Actions_Taken
  EditBox 75, 250, 90, 15, Worker_Signature
  ButtonGroup ButtonPressed
    OkButton 170, 250, 50, 15
    CancelButton 220, 250, 50, 15
  Text 5, 10, 50, 10, "Case Number:"
  Text 135, 10, 50, 10, "Completed By: "
  Text 5, 30, 90, 10, "Date Received By Agency: "
  GroupBox 165, 20, 100, 75, ""
  Text 170, 30, 90, 60, "RESOURCES: For more helpful information use the combined manual (footer month of 07/96) in MAXIS. Search for 'Foster Care', 'AFDC Assistance Standards', and/or 'IV-E'."
  Text 5, 50, 100, 10, "Date of Agency Responsiblity: "
  Text 5, 70, 20, 10, "IV E:"
  Text 5, 90, 45, 10, "SWKR or PO:"
  Text 5, 110, 25, 10, "AREP: "
  Text 5, 130, 30, 10, "Income: "
  Text 5, 150, 60, 10, "Retro Requested: "
  Text 5, 175, 30, 10, "HHcomp:"
  Text 10, 195, 20, 10, "OHC:"
  Text 5, 215, 80, 10, "Verifications Requested: "
  Text 10, 235, 50, 10, "Actions Taken:"
  Text 10, 255, 60, 10, "Worker Signature: "
  Text 125, 175, 30, 10, "Assets:"
EndDialog




'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------
'connecting to BlueZone, and grabbing the case number
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)

'calling the dialog---------------------------------------------------------------------------------------------------------------
DO
	Dialog Foster_Care_HCAPP
	IF buttonpressed = 0 THEN stopscript
	IF MAXIS_case_number = "" THEN MsgBox "You must have a case number to continue!"
	IF worker_signature = "" THEN MsgBox "You must enter a worker signature."
LOOP until MAXIS_case_number <> "" and worker_signature <> ""

'checking for an active MAXIS session
CALL check_for_MAXIS(FALSE)

'The case note---------------------------------------------------------------------------------------------------------------------
start_a_blank_CASE_NOTE
CALL write_variable_in_CASE_NOTE("***Foster Care HCAPP***")
CALL write_bullet_and_variable_in_CASE_NOTE("Date Received By Agency", Date_Received_By_Agency)
CALL write_bullet_and_variable_in_CASE_NOTE("Completed By", Completed_By)
CALL write_bullet_and_variable_in_CASE_NOTE("Date of Agency Responsibility", Date_Of_Agency_Responsiblity)
CALL write_bullet_and_variable_in_CASE_NOTE("IV E", IV_E)
CALL write_bullet_and_variable_in_CASE_NOTE("SWKR or PO", SWKR_or_PO)
CALL write_bullet_and_variable_in_CASE_NOTE("AREP", AREP)
CALL write_bullet_and_variable_in_CASE_NOTE("Income", Income)
CALL write_bullet_and_variable_in_CASE_NOTE("Retro Requested", Retro_Requested)
CALL write_bullet_and_variable_in_CASE_NOTE("HHcomp", HHComp)
CALL write_bullet_and_variable_in_CASE_NOTE("Assets", Assets)
CALL write_bullet_and_variable_in_CASE_NOTE("OHC", OHC)
CALL write_bullet_and_variable_in_CASE_NOTE("Verifications Requested", Verifications_Requested)
CALL write_bullet_and_variable_in_CASE_NOTE("Actions Taken", actions_taken)
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE(worker_signature)

Script_end_procedure("")
