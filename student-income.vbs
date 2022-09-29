'Required for statistical purposes===============================================================================
name_of_script = "DAIL - STUDENT INCOME.vbs"
start_time = timer
STATS_counter = 1              'sets the stats counter at one
STATS_manualtime = 64          'manual run time in seconds
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
'("11/28/2016", "Initial version.", "Charles Potter, DHS")
'END CHANGELOG BLOCK ======================================================================================================

EMConnect ""

'------------------THIS SCRIPT IS DESIGNED TO BE RUN FROM THE DAIL SCRUBBER.
'------------------As such, it does NOT include protections to be ran independently.

EMSendKey "n" + "<enter>"
EMWaitReady 1, 0
EMSendKey "<PF9>"
EMWaitReady 1, 0

EMSendKey "STUDENT INCOME HAS ENDED - REVIEW FS AND/OR HC RESULTS/APP" + "<newline>"
EMSendKey "* Sending financial aid form. TIKLed for 10-day return." + "<newline>"
EMSendKey "---" + "<newline>"

BeginDialog worker_sig_dialog, 0, 0, 191, 57, "Worker signature"
  EditBox 35, 25, 50, 15, worker_sig
  ButtonGroup ButtonPressed_worker_sig_dialog
    OkButton 135, 10, 50, 15
    CancelButton 135, 30, 50, 15
  Text 25, 10, 75, 10, "Sign your case note."
EndDialog

Dialog worker_sig_dialog
If ButtonPressed_worker_sig_dialog = 0 then stopscript

EMSendKey worker_sig + "<PF3>"
EMWaitReady 1, 0
EMSendKey "<PF3>"
EMWaitReady 1, 0

'Now it will TIKL out for this case.

EMSendKey "w" + "<enter>"
EMWaitReady 1, 0

'The following will generate a TIKL formatted date for 10 days from now.

If DatePart("d", Now + 10) = 1 then TIKL_day = "01"
If DatePart("d", Now + 10) = 2 then TIKL_day = "02"
If DatePart("d", Now + 10) = 3 then TIKL_day = "03"
If DatePart("d", Now + 10) = 4 then TIKL_day = "04"
If DatePart("d", Now + 10) = 5 then TIKL_day = "05"
If DatePart("d", Now + 10) = 6 then TIKL_day = "06"
If DatePart("d", Now + 10) = 7 then TIKL_day = "07"
If DatePart("d", Now + 10) = 8 then TIKL_day = "08"
If DatePart("d", Now + 10) = 9 then TIKL_day = "09"
If DatePart("d", Now + 10) > 9 then TIKL_day = DatePart("d", Now + 10)

If DatePart("m", Now + 10) = 1 then TIKL_month = "01"
If DatePart("m", Now + 10) = 2 then TIKL_month = "02"
If DatePart("m", Now + 10) = 3 then TIKL_month = "03"
If DatePart("m", Now + 10) = 4 then TIKL_month = "04"
If DatePart("m", Now + 10) = 5 then TIKL_month = "05"
If DatePart("m", Now + 10) = 6 then TIKL_month = "06"
If DatePart("m", Now + 10) = 7 then TIKL_month = "07"
If DatePart("m", Now + 10) = 8 then TIKL_month = "08"
If DatePart("m", Now + 10) = 9 then TIKL_month = "09"
If DatePart("m", Now + 10) > 9 then TIKL_month = DatePart("m", Now + 10)

TIKL_year = DatePart("yyyy", Now + 10)

EMSetCursor 5, 18
EMSendKey TIKL_month & TIKL_day & TIKL_year - 2000
EMSetCursor 9, 3
EMSendKey "Financial aid form should have returned by now. If not received and processed, take appropriate action. (TIKL auto-generated from script)." + "<enter>"
EMWaitReady 1, 0
EMSendKey "<PF3>"
MsgBox "MAXIS updated for student income status. A case note has been made, and a TIKL has been sent for 10 days from now. A financial aid form should be sent to the client."

script_end_procedure("")
