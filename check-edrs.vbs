'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - CHECK EDRS.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 49                	'manual run time in seconds
STATS_denomination = "M"       		'M is for each MEMBER
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

BeginDialog EDRS_dialog, 0, 0, 156, 80, "EDRS dialog"
  EditBox 60, 10, 80, 15, MAXIS_case_number
  ButtonGroup ButtonPressed
    OkButton 15, 55, 50, 15
    CancelButton 80, 55, 50, 15
  Text 5, 15, 50, 10, "Case Number:"
EndDialog



'THE SCRIPT----------------------------------------------------------------------------------------------------
EMConnect ""
'Hunts for Maxis case number to autofill it
Call MAXIS_case_number_finder(MAXIS_case_number)

'Error proof functions
Call check_for_MAXIS(true)

DO
	dialog EDRS_dialog
	IF buttonpressed = 0 THEN stopscript
	IF MAXIS_case_number = "" THEN MSGBOX "Please enter a case number"

LOOP UNTIL MAXIS_case_number <> ""

'Error proof functions
Call check_for_MAXIS(False)

'Creating a custom dialog for determining who the HH members are
call HH_member_custom_dialog(HH_member_array)

'Error proof functions
Call check_for_MAXIS(False)

'changing footer dates to current month to avoid invalid months.
MAXIS_footer_month = datepart("M", date)
	IF Len(MAXIS_footer_month) <> 2 THEN MAXIS_footer_month = "0" & MAXIS_footer_month
MAXIS_footer_year = right(datepart("YYYY", date), 2)

Dim Member_Info_Array()
Redim Member_Info_Array(UBound(HH_member_array), 4)


'Navigate to stat/memb and check for ERRR message
CALL navigate_to_MAXIS_screen("STAT", "MEMB")
For i = 0 to Ubound(HH_member_array)

	Member_Info_Array(i, 0) = HH_member_array(i)
	'Navigating to selected memb panel
	EMwritescreen HH_member_array(i), 20, 76
	transmit

	EMReadScreen no_MEMB, 13, 8, 22 'If this member does not exist, this will stop the script from continuing.
	IF no_MEMB = "Arrival Date:" THEN script_end_procedure("This HH member does not exist.")


	'Reading info and removing spaces
	EMReadscreen First_name, 12, 6, 63
	First_name = replace(First_name, "_", "")
	Member_Info_Array(i, 1) = First_name

	'Reading Last name and removing spaces
	EMReadscreen Last_name, 25, 6, 30
	Last_name = replace(Last_name, "_", "")
	Member_Info_Array(i, 2) = Last_name

	'Reading Middle initial and replacing _ with a blank if empty.
	EMReadscreen Middle_initial, 1, 6, 79
	Middle_initial = replace(Middle_initial, "_", "")
	Member_Info_Array(i, 3) = Middle_initial

	'Reads SSN
	Emreadscreen SSN_number, 11, 7, 42
	SSN_number = replace(SSN_number, " ", "")
	Member_Info_Array(i, 4) = SSN_number

	STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter

Next



'Navigate back to self and to EDRS
Back_to_self
CALL navigate_to_MAXIS_screen("INFC", "EDRS")
nameMatch = False
ssnMatch = False
Dim HitsArray()
ReDim HitsArray(UBound(HH_member_array)) 

For i = 0 to UBound(HH_member_array)

	'Write in SSN number into EDRS
	EMwritescreen Member_Info_Array(i, 4), 2, 7
	transmit
	Emreadscreen SSN_output, 7, 24, 2

	'Check to see what results you get from entering the SSN. If you get NO DISQ then check the person's name
	IF SSN_output = "NO DISQ" THEN
		EMWritescreen Member_Info_Array(i, 2), 2, 24
		EMWritescreen Member_Info_Array(i, 1), 2, 58
		EMWritescreen Member_Info_Array(i, 3), 2, 76
		transmit
		EMreadscreen NAME_output, 7, 24, 2
		IF NAME_output = "NO DISQ" THEN        'If after entering a name you still get NO DISQ then let worker know otherwise let them know you found a name.
			Hits = Hits & "No disqualifications found for Member #: " & Member_Info_Array(i, 0) & " " & Member_Info_Array(i, 1) & " " & Member_Info_Array(i, 2) & vbcr
			HitsArray(i) = "No disqualifications found for Member #: " & Member_Info_Array(i, 0) & " " & Member_Info_Array(i, 1) & " " & Member_Info_Array(i, 2)
		ELSE
			Hits = Hits & "Member #: " & Member_Info_Array(i, 0) & " " & Member_Info_Array(i, 1) & " " & Member_Info_Array(i, 2) & " has a potential name match. " & vbCr
			nameMatch = True
		END IF
	ELSE
		Hits = Hits & "Member #: " & Member_Info_Array(i, 0) & " " & Member_Info_Array(i, 1) & " " & Member_Info_Array(i, 2) & " has SSN Match. " & vbCr     'If after searching a SSN number you don't get the NO DISQ message then let worker know you found the SSN
		ssnMatch = True
		IF nameMatch	THEN 
			IF ssnMatch THEN
				matchString = "Member #: " & Member_Info_Array(i, 0) & " " & Member_Info_Array(i, 1) & " " & Member_Info_Array(i, 2) & " has both name and SSN Match"
			ELSE
				matchString = "Member #: " & Member_Info_Array(i, 0) & " " & Member_Info_Array(i, 1) & " " & Member_Info_Array(i, 2) & " has name Match."
		END IF
		ELSE matchString = "Member #: " & Member_Info_Array(i, 0) & " " & Member_Info_Array(i, 1) & " " & Member_Info_Array(i, 2) & " has SSN Match."
		HitsArray(i) = matchString
		END IF
	END IF
Next


'DIALOG-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
BeginDialog case_note_dialog, 0, 0, 270, 150, "Case note dialog"
  Text 10, 10, 260, 100, Hits
  CheckBox 10, 85, 260, 20, "Add Case Note with the above information", disqualification_checkbox
  Text 10, 105, 60, 10, "Worker signature:"
  EditBox 70, 105, 50, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 35, 125, 50, 15
    CancelButton 95, 125, 50, 15
EndDialog

disqualification_checkbox = checked

Do
	Do
		err_msg = ""
		Dialog case_note_dialog		'Displays the dialog
		cancel_confirmation
		If worker_signature = "" and disqualification_checkbox = checked THEN err_msg = err_msg & vbCr & "Please enter a worker signature, or Case Note manually."
		If err_msg <> "" THEN Msgbox err_msg
	Loop until (ButtonPressed = -1 and err_msg = "") 
call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = false
			
IF disqualification_checkbox = checked THEN 
	'Navigates to case note, and checks to make sure we aren't in inquiry.
	Back_to_self
	Call start_a_blank_CASE_NOTE
	noteHeader = "***** Check EDRS script ran "
	IF nameMatch = True THEN noteHeader = noteHeader + "- potential Name match found "
	IF ssnMatch = True THEN noteHeader = noteHeader + "- potential SSN match found "
	CALL write_variable_in_CASE_NOTE(noteHeader)
	For i = 0 to Ubound(HitsArray)
		CALL write_variable_in_CASE_NOTE(HitsArray(i))
	NEXT
	CALL write_variable_in_CASE_NOTE("---")
	CALL write_variable_in_CASE_NOTE(worker_signature)
END IF
STATS_counter = STATS_counter - 1			'Removing one instance of the STATS Counter
script_end_procedure("")
