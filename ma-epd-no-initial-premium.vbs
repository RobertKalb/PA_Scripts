'This script was developed by Charles Potter & Robert Kalb from Anoka County

'Required for statistical purposes==========================================================================================
name_of_script = "NOTICES - MA-EPD NO INITIAL PREMIUM.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 180                               'manual run time in seconds
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

'--- DIALOG-----------------------------------------------------------------------------------------------------------------------
BeginDialog WCOM_dlg, 0, 0, 196, 100, "MA-EPD No premium paid WCOM"
  EditBox 70, 15, 60, 15, MAXIS_case_number
  EditBox 70, 35, 30, 15, approval_month
  EditBox 160, 35, 30, 15, approval_year
  EditBox 80, 55, 60, 15, worker_signature
  Text 10, 20, 55, 10, "Case Number: "
  Text 10, 40, 55, 10, "Approval Month:"
  Text 105, 40, 55, 10, "Approval Year:"
  Text 10, 60, 70, 10, "Worker signature: "
  ButtonGroup ButtonPressed
    OkButton 50, 80, 50, 15
    CancelButton 105, 80, 50, 15
EndDialog
'--------------------------------------------------------------------------------------------------------------------------------

'--- The script -----------------------------------------------------------------------------------------------------------------
EMConnect ""

DO
	err_msg = ""
	dialog WCOM_dlg
	cancel_confirmation
	IF MAXIS_case_number = "" THEN err_msg = "Please enter a case number" & vbNewLine
	IF len(approval_month) <> 2 THEN err_msg = err_msg & "Please enter your month in MM format." & vbNewLine
	IF len(approval_year) <> 2 THEN err_msg = err_msg & "Please enter your year in YY format." & vbNewLine
	IF worker_signature = "" THEN err_msg = err_msg & "Please enter your worker signature." & vbNewLine
	IF err_msg <> "" THEN msgbox err_msg
LOOP until err_msg = ""

check_for_maxis(true)

CALL HH_member_custom_dialog(HH_member_array)

check_for_maxis(true)

call navigate_to_MAXIS_screen("spec", "wcom")

EMWriteScreen approval_month, 3, 46
EMWriteScreen approval_year, 3, 51
EMWriteScreen "Y", 3, 74 'selects HC only
transmit

'array created in previous menu selection
FOR each HH_member in HH_member_array
	DO 								'This DO/LOOP resets to the first page of notices in SPEC/WCOM
		EMReadScreen more_pages, 8, 18, 72
		IF more_pages = "MORE:  -" THEN PF7
	LOOP until more_pages <> "MORE:  -"

	read_row = 7
	DO
		waiting_check = ""
		EMReadscreen reference_number, 2, read_row, 62
		EMReadscreen waiting_check, 7, read_row, 71 'finds if notice has been printed
		If waiting_check = "Waiting" and reference_number = HH_member THEN 'checking program type and if it's been printed
			EMSetcursor read_row, 13
			EMSendKey "x"
			Transmit
			pf9
		    EMSetCursor 03, 15
      		EMWriteScreen "You are denied eligibility under Medical Assistance for", 3, 15
	      	EMWriteScreen "Employed Persons with Disabilities (MA-EPD) program because", 4, 15
	      	EMWriteScreen "the required premium was not paid by the due date, You may", 5, 15
			EMWriteScreen "request 'Good Cause' for late premium payment. This must be", 6, 15
			EMWriteScreen "approved by the Department of Human Services (DHS). To ", 7, 15
			EMWriteScreen "claim Good Cause, send a letter with your name, address,", 8, 15
			EMWriteScreen "case number and the reason for late payment to:", 9, 15
			EMWriteScreen "DHS MA-EPD Good Cause", 11, 15
			EMWriteScreen "P.O. Box 64967", 12, 15
			EMWriteScreen "St Paul, MN 55164-0967", 13, 15
			EMWriteScreen "Fax: 651 431 7563", 15, 15
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
	LOOP until reference_number = "  "
NEXT

If WCOM_count = 0 THEN
	MSGbox "No Waiting HC elig results were found in this month for this HH members."
	Stopscript
ELSE
	MSGbox "Success! A WCOM has been added."
END IF

script_end_procedure("")
