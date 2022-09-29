'Required for statistical purposes===============================================================================
name_of_script = "DAIL - FMED DEDUCTION.vbs"
start_time = timer
STATS_counter = 1              'sets the stats counter at one
STATS_manualtime = 127         'manual run time in seconds
STATS_denomination = "C"       'C is for case
'END OF stats block==============================================================================================

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
'("1/2/2018", "Fixing bug that prevented the script from writing SPEC/MEMO due to MAXIS updates. Additional updates to update syntax.", "Casey Love, Ramsey County")
'("11/28/2016", "Initial version.", "Charles Potter, DHS")
'END CHANGELOG BLOCK ======================================================================================================

'<<<<<GO THROUGH THE SCRIPT AND REMOVE REDUNDANT FUNCTIONS, THANKS TO CUSTOM FUNCTIONS THEY ARE NOT REQUIRED.

EMConnect ""

BeginDialog worker_sig_dialog, 0, 0, 141, 46, "Worker signature"
  EditBox 15, 25, 50, 15, worker_sig
  ButtonGroup ButtonPressed_worker_sig_dialog
    OkButton 85, 5, 50, 15
    CancelButton 85, 25, 50, 15
  Text 5, 10, 75, 10, "Sign your case note."
EndDialog

Dialog worker_sig_dialog
If ButtonPressed_worker_sig_dialog = 0 then stopscript

EMReadScreen MAXIS_case_number, 8, 5, 73
MAXIS_case_number = trim(MAXIS_case_number)

EMWriteScreen "P", 6, 3
transmit

EMWriteScreen "MEMO", 20, 70

start_a_new_spec_memo

Call write_variable_in_SPEC_MEMO ("You are turning 60 next month, so you may be eligible for a new deduction for SNAP. Clients who are over 60 years old may receive increased SNAP benefits if they have recurring medical bills over $35 each month.")
Call write_variable_in_SPEC_MEMO ("---")
Call write_variable_in_SPEC_MEMO ("If you have medical bills over $35 each month, please contact your worker to discuss adjusting your benefits. You will need to send in proof of the medical bills, such as pharmacy receipts, an explanation of benefits, or premium notices.")
Call write_variable_in_SPEC_MEMO ("  ")
Call write_variable_in_SPEC_MEMO ("Please call your worker with questions.")

PF4

EMWriteScreen "case", 19, 22
EMWriteScreen "note", 19, 70
transmit

start_a_blank_CASE_NOTE

Call write_variable_in_CASE_NOTE ("MEMBER HAS TURNED 60 - NOTIFY ABOUT POSSIBLE FMED DEDUCTION")
Call write_variable_in_CASE_NOTE ("---")
Call write_variable_in_CASE_NOTE ("* Sent MEMO to client about FMED deductions.")
Call write_variable_in_CASE_NOTE ("---")
Call write_variable_in_CASE_NOTE (worker_sig & ", using automated script.")

PF3

PF3

Call navigate_to_MAXIS_screen ("DAIL", "DAIL")

script_end_procedure("Success! The script has sent a MEMO to the client about the possible FMED deduction, and case noted the action.")
