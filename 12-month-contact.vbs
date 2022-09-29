'Required for statistical purposes==========================================================================================
name_of_script = "NOTICES - 12 MO CONTACT.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 90                               'manual run time in seconds
STATS_denomination = "C"       'C is for each CASE
'END OF stats block=========================================================================================================

'Because we are running these locally, we are going to get rid of all the calls to GitHub...
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
' 
' 'CHANGELOG BLOCK ===========================================================================================================
' 'Starts by defining a changelog array
' changelog = array()
' 
' 'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
' 'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
' call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
' changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DIALOG
BeginDialog case_number_dialog, 0, 0, 161, 61, "Case number"
  Text 5, 5, 85, 10, "Enter your case number:"
  EditBox 95, 0, 60, 15, MAXIS_case_number
  Text 5, 25, 70, 10, "Sign your case note:"
  EditBox 80, 20, 75, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 25, 40, 50, 15
    CancelButton 85, 40, 50, 15
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------
'grabbing case number & connecting to MAXIS
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)

dialog case_number_dialog
cancel_confirmation

'checking for an active MAXIS session
Call check_for_MAXIS(True)

'THE MEMO----------------------------------------------------------------------------------------------------
call navigate_to_MAXIS_screen("spec", "memo")
PF5
EMReadScreen MEMO_edit_mode_check, 26, 2, 28
If MEMO_edit_mode_check <> "Notice Recipient Selection" then
  MsgBox "You do not appear to be able to make a MEMO for this case. Are you in inquiry? Is this case out of county? Check these items and try again."
  Stopscript
End if
EMWriteScreen "x", 5, 12
transmit
Call write_variable_in_SPEC_MEMO ("************************************************************")
Call write_variable_in_SPEC_MEMO ("This notice is to remind you to report changes to your county worker by the 10th of the month following the month of the change. Changes that must be reported are address, people in your household, income, shelter costs and other changes such as legal obligation to pay child support. If you don't know whether to report a change, contact your county worker.")
Call write_variable_in_SPEC_MEMO ("************************************************************")
PF4

'THE CASE NOTE
call navigate_to_MAXIS_screen("case", "note")
PF9
Call write_variable_in_CASE_NOTE("Sent 12 month contact letter via SPEC/MEMO on " & date & ". -" & worker_signature)

script_end_procedure("")
