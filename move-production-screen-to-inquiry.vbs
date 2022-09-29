'Required for statistical purposes===============================================================================
name_of_script = "UTILITIES - MOVE PRODUCTION SCREEN TO INQUIRY.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 40                      'manual run time in seconds
STATS_denomination = "C"                   'C is for each CASE
'END OF stats block==============================================================================================

'Because we are running these locally, we are going to get rid of all the calls to GitHub...
FuncLib_URL = "I:\Blue Zone Scripts\Functions Library.vbs"
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

EMConnect "A"
row = 1
col = 1
EMSearch "Function: ", row, col
If row = 0 then
  MsgBox "Function not found."
  StopScript
End if
EMReadScreen MAXIS_function, 4, row, col + 10
If MAXIS_function = "____" then
  MsgBox "Function not found."
  StopScript
End if

row = 1
col = 1
EMSearch "Case Nbr: ", row, col
If row = 0 then
  MsgBox "Case number not found."
  StopScript
End if
EMReadScreen MAXIS_case_number, 8, row, col + 10

row = 1
col = 1
EMSearch "Month: ", row, col
If row = 0 then
  MsgBox "Footer month not found."
  StopScript
End if
EMReadScreen MAXIS_footer_month, 2, row, col + 7
EMReadScreen MAXIS_footer_year, 2, row, col + 10

row = 1
col = 1
EMSearch "(", row, col
If row = 0 then
  MsgBox "Command not found."
  StopScript
End if
EMReadScreen MAXIS_command, 4, row, col + 1
If MAXIS_command = "NOTE" then MAXIS_function = "CASE"

EMConnect "B"
EMFocus

attn
EMReadScreen inquiry_check, 7, 7, 15
If inquiry_check <> "RUNNING" then
  MsgBox "Inquiry not found. The script will now stop."
  StopScript
End if

EMWriteScreen "FMPI", 2, 15
transmit

back_to_self

EMWriteScreen MAXIS_function, 16, 43
EMWriteScreen "________", 18, 43
EMWriteScreen MAXIS_case_number, 18, 43
EMWriteScreen MAXIS_footer_month, 20, 43
EMWriteScreen MAXIS_footer_year, 20, 46
EMWriteScreen MAXIS_command, 21, 70
transmit

script_end_procedure("")
