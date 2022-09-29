'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - FOSTER CARE REVIEW.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 90           'manual run time in seconds
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
BeginDialog FC_HC_review_dialog, 0, 0, 256, 250, "FOSTER CARE HC REVIEW"
  EditBox 65, 5, 65, 15, MAXIS_case_number
  EditBox 65, 25, 65, 15, Received
  EditBox 65, 45, 65, 15, Completed_By
  EditBox 130, 70, 105, 15, Social_Worker_or_Probation_Officer
  EditBox 105, 90, 85, 15, Extended_Foster_Care_Date
  EditBox 40, 125, 55, 15, Income
  EditBox 40, 145, 70, 15, Results
  EditBox 75, 190, 110, 15, Worker_Signature
  ButtonGroup ButtonPressed
    OkButton 125, 230, 50, 15
    CancelButton 190, 230, 50, 15
  Text 5, 5, 55, 10, "Case number:"
  Text 5, 25, 45, 10, "Received: "
  Text 5, 45, 60, 10, "Completed By: "
  Text 5, 70, 120, 10, "Social Worker or Probation Officer:"
  Text 5, 95, 95, 15, "Extended Foster Care Date: "
  Text 5, 125, 35, 15, "Income: "
  Text 5, 150, 30, 15, "Results:"
  Text 5, 190, 65, 10, "Worker Signature: "
EndDialog



'THE SCRIPT--------------------------------------------------------------------------------------------------------------------------------------
'connecting to BlueZone, and grabbing the case number
EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)

'calling the dialog---------------------------------------------------------------------------------------------------------------
DO
	Dialog FC_HC_review_dialog
	IF buttonpressed = 0 THEN stopscript
	IF MAXIS_case_number = "" THEN MsgBox "You must have a case number to continue!"
	IF worker_signature = "" THEN MsgBox "You must enter a worker signature."
LOOP until MAXIS_case_number <> "" and worker_signature <> ""

'checking for an active MAXIS session
CALL check_for_MAXIS(FALSE)

'The case note---------------------------------------------------------------------------------------------------------------------
start_a_blank_CASE_NOTE
CALL write_variable_in_CASE_NOTE("***Foster Care HC REVIEW***")
CALL write_bullet_and_variable_in_CASE_NOTE("Received", Received)
CALL write_bullet_and_variable_in_CASE_NOTE("Completed By", Completed_By)
CALL write_bullet_and_variable_in_CASE_NOTE("Social Worker or Probation Officer", Social_Worker_or_Probation_Officer)
CALL write_bullet_and_variable_in_CASE_NOTE("Extended Foster Care Date", Extended_Foster_Care_Date)
CALL write_bullet_and_variable_in_CASE_NOTE("Income", Income)
CALL write_bullet_and_variable_in_CASE_NOTE("Results", Results)
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE(worker_signature)
Script_end_procedure("")
