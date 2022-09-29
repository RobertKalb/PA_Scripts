'This script was developed by Charles Potter from Anoka County
'Required for statistical purposes==========================================================================================
name_of_script = "NOTICES - POSTPONED WREG VERIFS.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 90                               'manual run time in seconds
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
' 
' 'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
' 'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
' call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")
' 
' 'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
' changelog_display
' 'END CHANGELOG BLOCK =======================================================================================================

'--- DIALOGS-----------------------------------------------------------------------------------------------------------------------
BeginDialog case_number_dlg, 0, 0, 196, 85, "Postponed WREG WCOM"
  EditBox 70, 15, 60, 15, MAXIS_case_number
  EditBox 70, 35, 30, 15, approval_month
  EditBox 160, 35, 30, 15, approval_year
  ButtonGroup ButtonPressed
    OkButton 45, 60, 50, 15
    CancelButton 100, 60, 50, 15
  Text 105, 40, 55, 10, "Approval Year:"
  Text 10, 20, 55, 10, "Case Number: "
  Text 10, 40, 55, 10, "Approval Month:"
EndDialog

BeginDialog WCOM_dlg, 0, 0, 196, 115, "Postponed WREG WCOM"
  EditBox 75, 10, 60, 15, ten_day_cutoff
  EditBox 55, 30, 60, 15, closure_date
  EditBox 70, 50, 115, 15, verifications_needed
  EditBox 75, 70, 60, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 45, 95, 50, 15
    CancelButton 100, 95, 50, 15
  Text 5, 75, 70, 10, "Worker signature: "
  Text 5, 35, 45, 10, "Closure Date: "
  Text 5, 55, 60, 10, "Verifs Postponed: "
  Text 5, 15, 65, 10, "Next 10 day cutoff: "
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
	IF err_msg <> "" THEN msgbox err_msg
LOOP until err_msg = ""

call check_for_maxis(false)

'Creating HH member array-------------------------------------------------------------------------------------------------------------
DO							'Loops until worker selects only one HH member. At this time the script only handles one HH member due to grammar issues involving multiple members with different postponed WREG verifs.
	Msgbox "Select the HH member that has the postponed verification. If you have multiple HH members please process manually at this time."
	CALL HH_member_custom_dialog(HH_member_array)
	array_length = Ubound(HH_member_array)
LOOP until array_length = 0

call check_for_maxis(false)

'Gathering/formatting variables---------------------------------------------------------------------------------------------------------------------
back_to_self
EMWriteScreen approval_month, 20, 43
EMWriteScreen approval_year, 20, 46
CALL check_for_maxis(false)
CALL navigate_to_MAXIS_screen("STAT", "PROG")   'grabbing the app date to use later.
EMReadScreen app_date, 8, 10, 33
app_date = replace(app_date, " ", "/")
CALL navigate_to_MAXIS_screen("STAT", "MEMB")  'grabbing client's name
EMWriteScreen HH_member_array(0), 20, 76
Transmit
EMReadScreen First_name, 12, 6, 63
EMReadScreen Last_name, 25, 6, 30
EMReadScreen Middle_initial, 1, 6, 79
client_name = replace(First_name, "_", "") & " " & replace(Middle_initial, "_", "") & " " & replace(Last_name, "_", "")
mid_month = left(app_date, 2) & "/15/" & right(app_date, 2)  'determining if the application date is before or after the 15th this affects when case must close by.
mid_month_check = datediff("D", mid_month, app_date)		'subtracting the day part of the app_date against the 15th
IF mid_month_check <= 0 THEN
	closure_date = dateadd("D", -1, (dateadd("M", 1, left(app_date, 2) & "/01/" & right(app_date, 2)))) ' 0 and below are before the 15th and close last day of the 1st month after application
	closure_date = cstr(closure_date)
	closure_month = right("00" & DatePart("M",closure_date), 2)				'defining the month to have 2 digits
	closure_year = right(closure_date, 2)									'creating a year variable to make the function parameters easier to read
	call ten_day_cutoff_check(closure_month, closure_year, ten_day_cutoff)
	ten_day_cutoff = cstr(ten_day_cutoff)
END IF
IF mid_month_check > 0 THEN
	closure_date = dateadd("D", -1, (dateadd("M", 2, left(app_date, 2) & "/01/" & right(app_date, 2))))	' anything above 0 is after the 15th and closes the last day of the 2nd month after application
	closure_date = cstr(closure_date)
	closure_month = right("00" & DatePart("M",closure_date), 2)				'defining the month to have 2 digits
	closure_year = right(closure_date, 2)									'creating a year variable to make the function parameters easier to read
	call ten_day_cutoff_check(closure_month, closure_year, ten_day_cutoff)
	ten_day_cutoff = cstr(ten_day_cutoff)
END IF

'2nd Dialog---------------------------------------------------------------------------------------------------------------------------------------------
DO
	err_msg = ""
	dialog WCOM_dlg
	cancel_confirmation
	IF ten_day_cutoff = "" THEN err_msg = err_msg & "Please enter your verif due date." & vbNewLine
	IF closure_date = "" THEN err_msg = err_msg & "Please enter your closure date." & vbNewLine
	IF verifications_needed = "" THEN err_msg = err_msg & "Please enter postponed verifications." & vbNewLine
	IF worker_signature = "" THEN err_msg = err_msg & "Please sign your NOTE" & vbNewLine
	IF err_msg <> "" THEN msgbox err_msg
LOOP until err_msg = ""

'error proofing as we don't want the script creating notices that have conflicting information on them.
If datediff("D", verif_due_date, closure_date) < 0 THEN script_end_procedure("Your verif due date is after your entered closure date. If this is the case please check policy and compose the WCOM manually.")

call check_for_maxis(false)

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
    	CALL write_variable_in_SPEC_MEMO(client_name & " has used their 3 entitled months of SNAP benefits as an Able Bodied Adult Without Dependents. ")
		CALL write_variable_in_SPEC_MEMO("Verification of " & verifications_needed & " has been postponed. You must turn in verification of " & verifications_needed & " by " & closure_date & " to continue to be eligible for SNAP benefits. ")
		CALL write_variable_in_SPEC_MEMO("If you do not turn in the required verifications, your case will close on " & closure_date & ".")
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
	'Case note
	start_a_blank_case_note
	call write_variable_in_CASE_NOTE("---WCOM regarding Postponed WREG verifs added---")
	call write_variable_in_CASE_NOTE("Client has had verification of " & verifications_needed & " postponed.")
	call write_bullet_and_variable_in_CASE_note("Next Cutoff Date", ten_day_cutoff)
	call write_bullet_and_variable_in_CASE_note("Closure Date", closure_date)
	call write_variable_in_CASE_NOTE("* TIKLed for next cutoff date. If verifs not received by cutoff date take appropriate action..")
	call write_variable_in_CASE_NOTE("---")
	call write_variable_in_CASE_NOTE(worker_signature)

	'Navigating to DAIL/WRIT
	call navigate_to_MAXIS_screen("dail", "writ")
	'The following will generate a TIKL formatted date for 10 days from now.
	call create_MAXIS_friendly_date(ten_day_cutoff, 0, 5, 18)
	'Writing in the rest of the TIKL.
	call write_variable_in_TIKL("Verification of postponed WREG verification(s) should have returned by now. If not received and processed, take appropriate action. (TIKL auto-generated from script)." )
	transmit
	PF3
	script_end_procedure("Success! The WCOM/CASE NOTE/TIKL have been added.")
END IF

script_end_procedure("")
