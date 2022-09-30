'Required for statistical purposes==========================================================================================
name_of_script = "ACTIONS - HOMELESS SHELTER DEDUCTION.vbs"
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

'Determines CM and CM+1 month and year using the two rightmost chars of both the month and year. Adds a "0" to all months,
' which will only pull over if it's a single-digit-month
Dim CM_mo, CM_yr, target_mo, target_yr
'var equals...  the right part of...    the specific part...    of either today or next month... just the right 2 chars!
CM_mo =         right("0" &             DatePart("m",           date                             ), 2)
CM_yr =         right(                  DatePart("yyyy",        date                             ), 2)

'Script to be released for use in October 
target_mo = "02"
target_yr= "23"

IF (Val(target_yr) >= Val(CM_yr)) THEN
	If Val(CM_yr) < Val(target_yr) THEN
 		script_end_procedure("This script is not available until 02/23")
 	end if
	If Val(CM_mo) < Val(target_mo) THEN
 		script_end_procedure("This script is not available until 02/23")
 	end if
end if

BeginDialog shelter_deduction_dialog, 0, 0, 156, 80, "Shelter Deduction dialog"
  EditBox 65, 10, 80, 15, MAXIS_case_number
  EditBox 85, 35, 20, 15, MAXIS_footer_month
  EditBox 110, 35, 20, 15, MAXIS_footer_year

  ButtonGroup ButtonPressed
    OkButton 15, 55, 50, 15
    CancelButton 80, 55, 50, 15
  Text 5, 15, 50, 10, "Case Number:"
  Text 5, 35, 65, 10, "Footer month/year:"
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------
EMConnect ""
'Hunts for Maxis case number and footer month/year to autofill
Call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

DO
	dialog shelter_deduction_dialog
	IF buttonpressed = 0 THEN stopscript
	IF MAXIS_case_number = "" THEN MSGBOX "Please enter a case number"

LOOP UNTIL MAXIS_case_number <> ""

'Creating a custom dialog for determining who the HH members are
call HH_member_custom_dialog(HH_member_array)

'Navigate to STAT/ADDR and check for homelessness
CALL navigate_to_MAXIS_screen("STAT", "ADDR")
EMReadScreen homeless_YN, 1, 10, 43
IF homeless_YN <> "Y" THEN script_end_procedure("This client is not homeless, the script will now close")

'navigate to ELIG/FS and check for an unapproved version
CALL navigate_to_MAXIS_screen("ELIG", "FS")
EMReadScreen approved_YN, 10, 3, 3
IF approved_YN <> "UNAPPROVED" THEN script_end_procedure("Unapproved version of SNAP does not exist. Please update the case and try again")

'Check FSB2 adjusted shelter cost 
CALL navigate_to_MAXIS_screen("ELIG", "FSB2")

'Reading income and adjusted shelter costs and converting them to decimals. Then checks if adj. shelter cost is too high for this to be helpful
EMReadScreen client_income, 7, 7, 73
EMReadScreen shelter_cost, 7, 17, 30
adjusted_income = Val(client_income)
adjusted_shelter_cost = Val(shelter_cost)
IF adjusted_shelter_cost > 166.80 THEN script_end_procedure("It may be more beneficial for the client to use the actual SHEL and HEST amounts. If the net income, SHEL amount, or HEST amounts are incorrect, please update the case and try again.")

Back_to_self
CALL navigate_to_MAXIS_screen("STAT","SHEL")

'Calculate relevant details for the deduction
calculated_shelter_cost = Round(adjusted_income/2, 2)
string_deduct = "" & calculated_shelter_cost
if right(string_deduct ,2) = ".5" Then
	calculated_shelter_cost = Val( (adjusted_income+1) / 2.0)
Else 
	calculated_shelter_cost = Val(adjusted_income / 2.0)
End if
calculated_shelter_cost = Round(calculated_shelter_cost, 0)
final_shel_deduction = Val(calculated_shelter_cost + 166.81)

'Update SHEL panel and delete HEST panel
'This is a dialog asking if the user wants to proceed with making changes.
    BeginDialog proceed_with_changes_dialog, 0, 0, 156, 80, "Proceed with changes dialog"
	Text 5, 10, 120, 50, "The script will now update the SHEL panel with the following deduction and delete the HEST panel, click OK to proceed: " & final_shel_deduction 
	ButtonGroup ButtonPressed
    	OkButton 15, 65, 50, 15
    	CancelButton 80, 65, 50, 15
	Text 10, 45, 60, 10, "Worker signature:"
	EditBox 70, 45, 50, 15, worker_signature
    EndDialog
Do
	Do
		err_msg = ""
		Dialog proceed_with_changes_dialog		'Displays the dialog
		cancel_confirmation
		If worker_signature = ""  THEN err_msg = err_msg & vbCr & "Please enter a worker signature for your Case Note."
		If err_msg <> "" THEN Msgbox err_msg
	Loop until (ButtonPressed = -1 and err_msg = "") 
call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = false

'Update SHEL panel
PF9
EMReadScreen does_shel_exist, len("PF9 IS NOT PERMITTED"), 24, 2
if does_shel_exist = "PF9 IS NOT PERMITTED" THEN 
	EMWriteScreen "nn", 20, 79 
	Transmit
end if
'HUD Subsizied(Y/N) set to NO 
EMWriteScreen "N", 6,46
'Shared (Y/N) set to NO
EMWriteScreen "N", 6, 64
'Paid to set to "Homeless Shel Deduction"
EMWriteScreen "  ", 7, 46
EMWriteScreen "                         ", 7, 50
EMWriteScreen "Homeless Shel Deduction", 7,50
'Rent Retrospective and Prospective set to calculated value, and verification fields set to OT
EMWriteScreen final_shel_deduction, 11,37
EMWriteScreen "OT", 11, 48
EMWriteScreen final_shel_deduction, 11,56
EMWriteScreen "OT", 11, 67
'All other Retrospective and Prospective Amount and Ver fields wiped out. 
FOR i = 12 To 18
	EMWriteScreen "        ", i, 37
	EMWriteScreen "  ", i, 48
	EMWriteScreen "        ", i, 56
	EMWriteScreen "  ", i, 67
Next

Transmit
DO
	EMReadScreen warning_message, 2, 24, 2
	IF warning_message <> "  " THEN Transmit
LOOP UNTIL warning_message = "  "

'Delete existing HEST panel
EMWriteScreen "HEST", 20, 71
Transmit
PF9
EMWriteScreen "DEL", 20, 71
Transmit

'Write CASE/NOTE for documentation
Back_to_self
CALL start_a_blank_CASE_NOTE
noteHeader = "// Homeless Shelter Deduction //"
CALL write_variable_in_CASE_NOTE(noteHeader)
noteAddLine = "*"
CALL write_variable_in_CASE_NOTE(noteAddLine)
noteLine2 = "Unit's net Income: $" & adjusted_income 
CALL write_variable_in_CASE_NOTE(noteLine2)
noteLinen = "Unit's income, halved: $" & string_deduct 
CALL write_variable_in_CASE_NOTE(noteLinen)
noteLine3 = "50% of the unit's net income: $" & adjusted_income & " / 2 = $" & calculated_shelter_cost
CALL write_variable_in_CASE_NOTE(noteLine3)
noteLine4 = "Homeless Deduction Amount entered on SHEL: $" & adjusted_income & "/2+166.81 = $" & final_shel_deduction
CALL write_variable_in_CASE_NOTE(noteLine4)
noteLine5 = "This case was determined to benefit from the homeless shelter deduction amount as it was higher than the adjusted shelter costs reported by the client."
CALL write_variable_in_CASE_NOTE(noteLine5)
noteLine6 = "Utility deduction not allowed for case eligible for homeless deduction. STAT/HEST panel deleted."
CALL write_variable_in_CASE_NOTE(noteLine6)
noteLine7 = "Deduction starting: " & MAXIS_footer_month & "/" & MAXIS_footer_year
CALL write_variable_in_CASE_NOTE(noteLine7)
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE(worker_signature)

'Navigate to ELIG/FS and prompt for review
Back_to_self
CALL Navigate_to_MAXIS_screen("ELIG","FS")

STATS_counter = STATS_counter - 1			'Removing one instance of the STATS Counter

script_end_procedure("Please review eligibility results before approving")
