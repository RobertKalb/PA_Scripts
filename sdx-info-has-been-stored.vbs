'Required for statistical purposes===============================================================================
name_of_script = "DAIL - SDX INFO HAS BEEN STORED.vbs"
start_time = timer
STATS_counter = 1              'sets the stats counter at one
STATS_manualtime = 10          'manual run time in seconds
STATS_denomination = "C"       'C is for Case
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

'------------------THIS SCRIPT IS DESIGNED TO BE RUN FROM THE DAIL SCRUBBER.
'------------------As such, it does NOT include protections to be ran independently.

EMConnect ""
EMSendKey "i" + "<enter>"

EMWaitReady 0, 0
EMSetCursor 20, 71
EMSendKey "sdxs" + "<enter>"

EMWaitReady 0, 0

script_end_procedure("")
