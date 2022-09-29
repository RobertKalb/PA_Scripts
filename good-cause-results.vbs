'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - GOOD CAUSE RESULTS.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 320                     'manual run time in seconds
STATS_denomination = "C"                   'C is for each CASE
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

BeginDialog Good_Cause_Claimed_Results_Dialog, 0, 0, 276, 300, "Good Cause Claim Determination"
  EditBox 205, 20, 65, 15, MAXIS_case_number
  EditBox 135, 40, 60, 15, Claim_Committee_Date
  DropListBox 120, 60, 105, 15, "Select One:"+chr(9)+"APPROVED"+chr(9)+"DENIED", determination_droplist
  EditBox 100, 100, 60, 15, Approved_to_Date
  EditBox 105, 140, 60, 15, DHS3629_Sent_Date
  EditBox 85, 180, 165, 15, Denial_Reason
  EditBox 160, 200, 65, 15, Date_DHS_docs_sent
  CheckBox 120, 225, 30, 15, "CCAP", CCAP_CHECKBOX
  CheckBox 155, 225, 30, 15, "DWP", DWP_CHECKBOX
  CheckBox 190, 225, 50, 15, "Health Care", HC_CHECKBOX
  CheckBox 245, 225, 30, 15, "MFIP", MFIP_CHECKBOX
  EditBox 35, 250, 235, 15, Other_comments
  EditBox 75, 280, 75, 15, Worker_signature
  ButtonGroup ButtonPressed
    OkButton 165, 280, 50, 15
    CancelButton 220, 280, 50, 15
  Text 40, 5, 195, 10, "Child Support Good Cause Exemption Claim Determination"
  Text 150, 20, 50, 15, "Case Number"
  Text 5, 40, 130, 15, "Date of Good Cause committee review:"
  Text 5, 60, 110, 15, "Is the claim approved or denied?"
  GroupBox 0, 85, 265, 75, "IF APPROVED - COMPLETE THE FOLLOWING:"
  Text 15, 100, 80, 15, "Date approved through*:"
  Text 35, 120, 185, 15, "*NOTE: A TIKL will be set for the date entered."
  Text 15, 140, 85, 15, "Date DHS-3629 was sent:"
  GroupBox 0, 170, 265, 50, "IF DENIED - COMPLETE THE FOLLOWING:"
  Text 20, 180, 60, 15, "Reason for denial:"
  Text 20, 200, 135, 15, "Date DHS-3628 and DHS-0033 were sent:"
  Text 10, 225, 105, 15, "Programs - select all that apply:"
  Text 10, 255, 20, 15, "Other:"
  Text 10, 280, 60, 15, "Worker Signature"
EndDialog

'Script----------------------------------------------
'Connect to Bluezone
EMConnect ""

'Inserts Maxis Case number
CALL MAXIS_case_number_finder(MAXIS_case_number)

'Shows dialog

DO
	err_msg = ""
	Dialog Good_Cause_Claimed_Results_Dialog
	cancel_confirmation
	IF IsNumeric(MAXIS_case_number)=FALSE THEN err_Msg = err_msg & vbCr & "You must type a valid numeric case number."
	IF Determination_droplist = "Select One:" THEN err_Msg = err_msg & vbCr & "You must select Approved or Denied."
	IF (Determination_droplist = "APPROVED" AND isdate(Approved_to_date) = FALSE) THEN err_Msg = err_msg & vbCr & "DAIL/TIKL date is not a valid date, please use MM/DD/YYYY format."
	IF worker_signature = "" THEN err_Msg = err_msg & vbCr & "You must sign your case note!"
	IF err_msg <> "" THEN Msgbox err_msg
LOOP UNTIL err_msg = ""



'seting variables for the programs included
IF CCAP_checkbox = 1 THEN programs_included = programs_included & "CCAP "
IF DWP_checkbox = 1 THEN programs_included = programs_included & "DWP "
IF MFIP_checkbox = 1 THEN programs_included = programs_included & "MFIP "
IF HC_checkbox = 1 THEN programs_included = programs_included & "Healthcare "

'Checks Maxis for password prompt
CALL check_for_MAXIS(True)

'Navigates to case note
CALL start_a_blank_CASE_NOTE

'Writes the case note
CALL write_variable_in_case_note (">>Child Support Good Cause Exemption Claimed - Determination: " & determination_droplist & "<<")
CALL write_bullet_and_variable_in_case_note("The Good Cause Committee review was on", claim_committee_date)
IF Determination_droplist = "APPROVED" THEN CALL write_bullet_and_variable_in_case_note("Date approved through", approved_to_date & " - DAIL/TIKL was created for this date")
CALL write_bullet_and_variable_in_case_note("Reason for denial", denial_reason)
CALL write_bullet_and_variable_in_case_note("Applicable Programs", programs_included)
CALL write_bullet_and_variable_in_case_note("Date DHS-3629 was sent", dhs3629_sent_DATE)
CALL write_bullet_and_variable_in_case_note("Date DHS-3628 & DHS-0033 were sent", Date_DHS_docs_sent)
CALL write_bullet_and_variable_in_case_note("Additional information", Other_Comments)



CALL write_variable_in_case_note("---")
CALL write_variable_in_case_note(worker_signature)

'TIKL PROCESS for APPROVED claims only
If approved_to_date<> "" then
		back_to_self
		call navigate_to_MAXIS_screen("DAIL", "WRIT")
		call create_MAXIS_friendly_date(approved_to_date, 0, 5, 18)
		call write_variable_in_TIKL("Good Cause claim needs to be reviewed.")
		PF3
	End if



script_end_procedure("Success! A case note has been made.  If the Good Cause claim was approved, a TIKL was also made.")
