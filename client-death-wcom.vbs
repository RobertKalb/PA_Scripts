'Required for statistical purposes==========================================================================================
name_of_script = "NOTICES - CLIENT DEATH WCOM.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 60                               'manual run time in seconds
STATS_denomination = "C"       'C is for each CASE
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
' 
' 'CHANGELOG BLOCK ===========================================================================================================
' 'Starts by defining a changelog array
' changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
'call changelog_update("01/17/2017", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
'changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'Dialog--------------------------------------------
BeginDialog death_dlg, 0, 0, 156, 70, "Client Death WCOM"
  EditBox 65, 5, 75, 15, MAXIS_case_number
  EditBox 60, 25, 20, 15, MAXIS_footer_month
  EditBox 130, 25, 20, 15, MAXIS_footer_year
  ButtonGroup ButtonPressed
    OkButton 25, 50, 50, 15
    CancelButton 80, 50, 50, 15
  Text 10, 10, 55, 10, "Case Number: "
  Text 10, 25, 45, 20, "Footer Month (MM):"
  Text 85, 25, 40, 20, "Footer Year (YY):"
EndDialog

'The script-------------------------------------
EMConnect ""

'warning box
Msgbox "Warning: If you have multiple waiting SNAP results this script may be unable to find the most recent one. Please process manually in those instances."

'the dialog
Do
	Do
		Do
			dialog death_dlg
			cancel_confirmation
			If MAXIS_footer_month = "" or MAXIS_footer_year = "" THEN Msgbox "Please fill in footer month and year (MM YY format)."
			If MAXIS_case_number = "" THEN MsgBox "Please enter a case number."
			If worker_signature = "" THEN MsgBox "Please sign your note."
		Loop until MAXIS_footer_month <> "" & MAXIS_footer_year <> ""
	Loop until MAXIS_case_number <> ""
Loop until worker_signature <> ""

'Converting dates into useable forms
If len(MAXIS_footer_month) < 2 THEN MAXIS_footer_month = "0" & MAXIS_footer_month
If len(MAXIS_footer_year) > 2 THEN MAXIS_footer_year = right(MAXIS_footer_year, 2)

'Navigating to the spec wcom screen
CALL Check_for_MAXIS(false)

Emwritescreen MAXIS_case_number, 18, 43
Emwritescreen MAXIS_footer_month, 20, 43
Emwritescreen MAXIS_footer_year, 20, 46

CALL navigate_to_MAXIS_screen("SPEC", "WCOM")

'Searching for waiting SNAP notice
wcom_row = 6
Do
	wcom_row = wcom_row + 1
	Emreadscreen program_type, 2, wcom_row, 26
	Emreadscreen print_status, 7, wcom_row, 71
	If program_type = "FS" then
		If print_status = "Waiting" then
			Emwritescreen "x", wcom_row, 13
			Transmit
			PF9
			Emreadscreen fs_wcom_exists, 3, 3, 15
			If fs_wcom_exists <> "   " then script_end_procedure ("It appears you already have a WCOM added to this notice. The script will now end.")
			If program_type = "FS" AND print_status = "Waiting" then
				fs_wcom_writen = true
				'This will write if the notice is for SNAP only
				CALL write_variable_in_SPEC_MEMO("******************************************************")
				CALL write_variable_in_SPEC_MEMO("")
				CALL write_variable_in_SPEC_MEMO("This SNAP case has been closed because the only eligible unit member has died.")
				CALL write_variable_in_SPEC_MEMO("")
				CALL write_variable_in_SPEC_MEMO("******************************************************")
				PF4
				PF3
			End if
		End If
	End If
	If fs_wcom_writen = true then Exit Do
	If wcom_row = 17 then
		PF8
		Emreadscreen spec_edit_check, 6, 24, 2
		wcom_row = 6
	end if
	If spec_edit_check = "NOTICE" THEN no_fs_waiting = true
Loop until spec_edit_check = "NOTICE"

If no_fs_waiting = true AND no_mf_waiting = true then script_end_procedure("No waiting FS notice was found for the requested month")

script_end_procedure("WCOM has been added to the first found waiting SNAP notice for the month and case selected. Please review the notice.")
