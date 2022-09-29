EMConnect ""

'------------------THIS SCRIPT IS DESIGNED TO BE RUN FROM THE DAIL SCRUBBER.
'------------------As such, it does NOT include protections to be ran independently.

'Required for statistical purposes===============================================================================
name_of_script = "DAIL - AFFILIATED CASE LOOKUP.vbs"
start_time = timer
STATS_counter = 1              'sets the stats counter at one
STATS_manualtime = 10          'manual run time in seconds
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

  row = 1
  col = 1
  cola = 1
EMSearch "#", 6, col
EMSearch ")", 6, cola

case_number_digits = cola - col - 1
EMReadScreen MAXIS_case_number, case_number_digits, 6, col + 1
If IsNumeric(MAXIS_case_number) = False then MsgBox "An affiliated case could not be detected on this dail message. Try another dail message."
If IsNumeric(MAXIS_case_number) = False then stopscript

'This Do...loop gets back to SELF.
Do
     EMWaitReady 1, 0
     EMReadScreen SELF_check, 27, 2, 28
     If SELF_check <> "Select Function Menu (SELF)" then EMSendKey "<PF3>"
Loop until SELF_check = "Select Function Menu (SELF)"


EMSetCursor 16, 43
EMSendKey "case"
EMSetCursor 18, 43
EMSendKey "<eraseEOF>" + MAXIS_case_number
EMSetCursor 21, 70
EMSendKey "note" + "<enter>"

MsgBox "You are now in case/note for the affiliated case!"

script_end_procedure("")
