'Required for statistical purposes==========================================================================================
name_of_script = "NOTICES - ABAWD WITH CHILD IN HH WCOM.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 60                               'manual run time in seconds
STATS_denomination = "C"       'C is for each CASE
'END OF stats block==============================================================================================

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
' 
' 'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
' 'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
' call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
' changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'--- DIALOGS-----------------------------------------------------------------------------------------------------------------------
BeginDialog case_number_dlg, 0, 0, 196, 100, "Adding ABAWD Adult WCOM"
  EditBox 70, 15, 60, 15, MAXIS_case_number
  EditBox 70, 35, 30, 15, approval_month
  EditBox 160, 35, 30, 15, approval_year
  EditBox 95, 55, 60, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 45, 80, 50, 15
    CancelButton 100, 80, 50, 15
  Text 105, 40, 55, 10, "Approval Year:"
  Text 10, 20, 55, 10, "Case Number: "
  Text 10, 40, 55, 10, "Approval Month:"
  Text 25, 60, 70, 10, "Worker signature: "
EndDialog

'--------------------------------------------------------------------------------------------------------------------------------

'--- The script -----------------------------------------------------------------------------------------------------------------

EMConnect ""

call MAXIS_case_number_finder(MAXIS_case_number)

'1st Dialog ---------------------------------------------------------------------------------------------------------------------
DO
	err_msg = ""
	dialog case_number_dlg
	cancel_confirmation
	IF MAXIS_case_number = "" THEN err_msg = "Please enter a case number" & vbNewLine
	IF len(approval_month) <> 2 THEN err_msg = err_msg & "Please enter your month in MM format." & vbNewLine
	IF len(approval_year) <> 2 THEN err_msg = err_msg & "Please enter your year in YY format." & vbNewLine
	IF worker_signature = "" THEN err_msg = err_msg & "Please enter your worker signature." & vbNewLine
	IF err_msg <> "" THEN msgbox err_msg
LOOP until err_msg = ""

call check_for_maxis(false)

'Creating HH member array-------------------------------------------------------------------------------------------------------------
Msgbox "Select the SNAP ABAWD member(s) that is exempt due to a child under 18."
CALL HH_member_custom_dialog(HH_member_array)


call check_for_maxis(false)

'Gathering/formatting variables---------------------------------------------------------------------------------------------------------------------
back_to_self
EMWriteScreen approval_month, 20, 43
EMWriteScreen approval_year, 20, 46
CALL check_for_maxis(false)
FOR each HH_member in HH_member_array
	CALL navigate_to_MAXIS_screen("STAT", "MEMB")  'grabbing client's name
	EMWriteScreen HH_member, 20, 76
	Transmit
	EMReadScreen First_name, 12, 6, 63
	EMReadScreen Last_name, 25, 6, 30
	EMReadScreen Middle_initial, 1, 6, 79
	client_name = client_name & replace(First_name, "_", "") & " " & replace(Middle_initial, "_", "") & " " & replace(Last_name, "_", "") & ", "
NEXT

'WCOM PIECE---------------------------------------------------------------------------------------------------------------------
call navigate_to_MAXIS_screen("spec", "wcom")

EMWriteScreen approval_month, 3, 46
EMWriteScreen approval_year, 3, 51
transmit

DO 								'This DO/LOOP resets to the first page of notices in SPEC/WCOM
	EMReadScreen more_pages, 8, 18, 72
	IF more_pages = "MORE:  -" THEN PF7
LOOP until more_pages <> "MORE:  -"

read_row = 7
DO
	waiting_check = ""
	EMReadscreen prog_type, 2, read_row, 26
	EMReadscreen waiting_check, 7, read_row, 71 'finds if notice has been printed
	If waiting_check = "Waiting" and prog_type = "FS" THEN 'checking program type and if it's been printed
		EMSetcursor read_row, 13
		EMSendKey "x"
		Transmit
		pf9
	    EMSetCursor 03, 15
    	CALL write_variable_in_SPEC_MEMO(client_name & " is(are) exempt from Able Bodied Adult Without Dependents(ABAWD) work requirements due to a child(ren) under the age of 18 in the SNAP unit. ")
	    PF4
		PF3
		WCOM_count = WCOM_count + 1
		exit do
	ELSE
		read_row = read_row + 1
	END IF
	IF read_row = 18 THEN
		PF8          'Navigates to the next page of notices.  DO/LOOP until read_row = 18
		read_row = 7
	End if
LOOP until prog_type = "  "

'Outcome ---------------------------------------------------------------------------------------------------------------------

If WCOM_count = 0 THEN  'if no waiting FS notice is found
	script_end_procedure("No Waiting FS elig results were found in this month for this HH member.")
ELSE 					'If a waiting FS notice is found
	script_end_procedure("Success! The WCOM/CASE NOTE/TIKL have been added.")
END IF

script_end_procedure("")
