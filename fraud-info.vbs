'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - FRAUD INFO.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 120          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

'LOADING FUNCTIONS LIBRARY FROM REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		FuncLib_URL = script_repository & "MAXIS FUNCTIONS LIBRARY.vbs"
		critical_error_msgbox = MsgBox ("The Functions Library code was not able to be reached by " &name_of_script & vbNewLine & vbNewLine &_
                                            "FuncLib URL: " & FuncLib_URL & vbNewLine & vbNewLine &_
                                            "The script has stopped. Send issues to " & contact_admin , _
                                            vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
            StopScript
	ELSE
		FuncLib_URL = script_repository & "MAXIS FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'CHANGELOG BLOCK ===========================================================================================================
'("10/16/2019", "All infrastructure changed to run locally and stored in BlueZone Scripts ccm. MNIT @ DHS)
'("11/28/2016", "Initial version.", "Charles Potter, DHS")
'END CHANGELOG BLOCK =======================================================================================================

'Dialog---------------------------------------------------------------------------------------------------------------------------
BeginDialog Fraud_Dialog, 0, 0, 211, 275, "Fraud Info"
  EditBox 65, 10, 90, 15, MAXIS_case_number
  EditBox 75, 30, 115, 15, referral_date
  EditBox 10, 65, 195, 15, referral_reason
  EditBox 10, 100, 195, 15, fraud_findings
  EditBox 10, 135, 195, 15, actions_taken
  DropListBox 10, 170, 55, 15, "Select One..."+chr(9)+"Yes"+chr(9)+"No"+chr(9)+"TBD", overpayment_yn
  EditBox 100, 230, 95, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 100, 250, 50, 15
    CancelButton 155, 250, 50, 15
  Text 10, 15, 55, 10, "Case Number: "
  Text 10, 35, 65, 10, "Date referral made:"
  Text 10, 50, 110, 10, "Reason for referral (be specific):"
  Text 10, 85, 55, 10, "Fraud findings:"
  Text 10, 120, 50, 10, "Actions taken:"
  Text 10, 155, 50, 10, "Overpayment?"
  Text 10, 190, 90, 35, "If yes for overpayment please use overpayment script to case note the specific details regarding it. "
  Text 35, 235, 60, 10, "Worker Signature: "
  Text 120, 155, 85, 50, "NOTE: You can type ; to seperate text with a new line in the case note. EX. 'This client; will need' would be written in two lines"
EndDialog

'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------
'connecting to MAXIS session and finding case number
EMConnect ""
CALL MAXIS_case_number_finder(MAXIS_case_number)

'calling the dialog---------------------------------------------------------------------------------------------------------------
DO
	err_msg = ""
	Dialog fraud_dialog
	cancel_confirmation
	IF MAXIS_case_number = "" THEN err_msg = "You must have a case number to continue!"
	IF worker_signature = "" THEN err_msg = err_msg & vbNewLine & "You must enter a worker signature."
	IF overpayment_yn = "Select One..." THEN err_msg = err_msg & vbNewLine & "You must select an option for overpayment."
	IF err_msg <> "" THEN msgbox "*** Notice!!! ***" & vbNewLine & err_msg
LOOP until err_msg = ""

'checking for an active MAXIS session
CALL check_for_MAXIS(False)

IF overpayment_yn = "Yes" THEN overpayment_yn = " Yes. See overpayment case note for more details."

'The case note---------------------------------------------------------------------------------------------------------------------
start_a_blank_CASE_NOTE
CALL write_variable_in_CASE_NOTE("***Fraud Referral Info***")
CALL write_bullet_and_variable_in_CASE_NOTE("Referral Date", referral_date)
CALL write_bullet_and_variable_in_CASE_NOTE("Referral Reason", referral_reason)
CALL write_bullet_and_variable_in_CASE_NOTE("Findings", fraud_findings)
CALL write_bullet_and_variable_in_CASE_NOTE("Actions Taken", actions_taken)
CALL write_bullet_and_variable_in_CASE_NOTE("Overpayment?", overpayment_yn)
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE(worker_signature)

Script_end_procedure("")
