'Required for statistical purposes===============================================================================
name_of_script = "DAIL - FMED DEDUCTION.vbs"
start_time = timer
STATS_counter = 1              'sets the stats counter at one
STATS_manualtime = 127         'manual run time in seconds
STATS_denomination = "C"       'C is for case
'END OF stats block==============================================================================================

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

'<<<<<GO THROUGH THE SCRIPT AND REMOVE REDUNDANT FUNCTIONS, THANKS TO CUSTOM FUNCTIONS THEY ARE NOT REQUIRED.

EMConnect ""

Set objNet = CreateObject("WScript.NetWork") 
windows_user_ID = UCASE(objNet.UserName)


BeginDialog worker_sig_dialog, 0, 0, 141, 46, "Worker signature"
  EditBox 15, 25, 50, 15, worker_sig
  ButtonGroup ButtonPressed_worker_sig_dialog
    OkButton 85, 5, 50, 15
    CancelButton 85, 25, 50, 15
  Text 5, 10, 75, 10, "Sign your case note."
EndDialog

Dialog worker_sig_dialog
If ButtonPressed_worker_sig_dialog = 0 then stopscript

EMWriteScreen "p", 6, 3
TRANSMIT

EMWriteScreen "memo", 20, 70
TRANSMIT

PF5

EMWriteScreen "x", 5, 12
Transmit

EMSendKey "You are turning 60 next month, so you may be eligible for a new deduction for SNAP." + "<newline>" + "<newline>"
EMSendKey "Clients who are over 60 years old may receive increased SNAP benefits if they have recurring medical bills over $35 each month." + "<newline>" + "<newline>"
EMSendKey "If you have medical bills over $35 each month, please contact your worker to discuss adjusting your benefits. You will need to send in proof of the medical bills, such as pharmacy receipts, an explanation of benefits, or premium notices." + "<newline>" + "<newline>"
EMSendKey "Please call your worker with questions."

PF4

EMReadScreen maxis_case_number, 8, 19, 38
maxis_case_number = replace(maxis_case_number, " ", "")
CALL navigate_to_MAXIS_screen("CASE", "NOTE")

DO
	PF9
	EMReadScreen case_note_check, 17, 2, 33
	EMReadScreen mode_check, 1, 20, 09
	If case_note_check <> "Case Notes (NOTE)" or mode_check <> "A" then msgbox "The script can't open a case note. Reasons may include:" & vbnewline & vbnewline & "* You may be in inquiry" & vbnewline & "* You may not have authorization to case note this case (e.g.: out-of-county case)" & vbnewline & vbnewline & "Check MAXIS and/or navigate to CASE/NOTE, and try again. You can press the STOP SCRIPT button on the power pad to stop the script."
Loop until (mode_check = "A" or mode_check = "E")

call write_variable_in_CASE_NOTE("MEMBER HAS TURNED 60 - NOTIFY ABOUT POSSIBLE FMED DEDUCTION")
call write_variable_in_CASE_NOTE("---")
call write_variable_in_CASE_NOTE("* Sent MEMO to client about FMED deductions.")
call write_variable_in_CASE_NOTE("---")
call write_variable_in_CASE_NOTE(worker_sig)


PF3
PF3

MsgBox "The script has sent a MEMO to the client about the possible FMED deduction, and case noted the action."

script_end_procedure("")
