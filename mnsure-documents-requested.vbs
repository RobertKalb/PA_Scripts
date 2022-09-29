'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - MNSURE - DOCUMENTS REQUESTED.vbs"
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

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog MNsure_docs_reqd_dialog, 0, 0, 301, 105, "MNsure Docs Req'd Dialog"
  EditBox 75, 5, 70, 15, MAXIS_case_number
  EditBox 225, 5, 70, 15, MNsure_app_date
  EditBox 45, 25, 70, 15, MNsure_ID
  EditBox 225, 25, 70, 15, application_case_number
  EditBox 50, 45, 245, 15, docs_reqd
  EditBox 50, 65, 245, 15, other_notes
  EditBox 70, 85, 70, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 190, 85, 50, 15
    CancelButton 245, 85, 50, 15
  Text 5, 10, 70, 10, "MAXIS case number:"
  Text 165, 10, 60, 10, "MNsure app date:"
  Text 5, 30, 40, 10, "MNsure ID:"
  Text 140, 30, 85, 10, "Application Case Number:"
  Text 5, 50, 40, 10, "Doc's req'd:"
  Text 5, 70, 45, 10, "Other notes:"
  Text 5, 90, 60, 10, "Worker signature:"
EndDialog


'THE SCRIPT----------------------------------------------------------------------------------------------------
'connecting to MAXIS
EMConnect ""
'Finds the case number
call MAXIS_case_number_finder(MAXIS_case_number)

'Displays the dialog and navigates to case note
Do
	Dialog MNsure_docs_reqd_dialog
	cancel_confirmation
	If MAXIS_case_number = "" then MsgBox "You must have a case number to continue!"
Loop until MAXIS_case_number <> ""


'checking for an active MAXIS session
Call check_for_MAXIS(False)


'THE CASE NOTE----------------------------------------------------------------------------------------------------
Call start_a_blank_CASE_NOTE
Call write_variable_in_CASE_NOTE(">>>>>MNSURE DOCS REQ'D<<<<<")
If MNsure_app_date <> "" then call write_bullet_and_variable_in_case_note("MNsure application date", MNsure_app_date)
If MNsure_ID <> "" then call write_bullet_and_variable_in_case_note("MNsure ID", MNsure_ID)
If application_case_number <> "" then call write_bullet_and_variable_in_case_note("Application case number", application_case_number)
If docs_reqd <> "" then call write_bullet_and_variable_in_case_note("Docs requested", docs_reqd)
If other_notes <> "" then call write_bullet_and_variable_in_case_note("Other notes", other_notes)
call write_bullet_and_variable_in_case_note("Please note", "If these docs come into your ''My documents received'' queue in OnBase, please create a copy of the document and re-index it to the appropriate MNsure doc type, and send to the proper workflow. If you have questions, consult a member of the MNsure team.")
call write_variable_in_CASE_NOTE("---")
call write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure("")
