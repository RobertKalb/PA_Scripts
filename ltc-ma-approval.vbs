'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - LTC - MA APPROVAL.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 300          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
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
'END CHANGELOG BLOCK =======================================================================================================

'>>>>NOTE: these were added as a batch process. Check below for any 'StopScript' functions and convert manually to the script_end_procedure("") function

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog case_number_dialog, 0, 0, 161, 61, "Case number"
  Text 5, 5, 85, 10, "Enter your case number:"
  EditBox 95, 0, 60, 15, MAXIS_case_number
  Text 15, 25, 50, 10, "Footer month:"
  EditBox 65, 20, 25, 15, MAXIS_footer_month
  Text 95, 25, 20, 10, "Year:"
  EditBox 120, 20, 25, 15, MAXIS_footer_year
  ButtonGroup ButtonPressed
    OkButton 25, 40, 50, 15
    CancelButton 85, 40, 50, 15
EndDialog

BeginDialog BBUD_Dialog, 0, 0, 191, 76, "BBUD"
  Text 5, 10, 180, 10, "This is a method B budget. What would you like to do?"
  ButtonGroup ButtonPressed
    PushButton 20, 25, 70, 15, "Jump to STAT/BILS", BILS_button
    PushButton 100, 25, 70, 15, "Stay in ELIG/HC", ELIG_button
    CancelButton 135, 55, 50, 15
EndDialog

BeginDialog approval_dialog, 0, 0, 376, 165, "Approval dialog"
  DropListBox 45, 5, 30, 15, "AX"+chr(9)+"EX"+chr(9)+"DX"+chr(9)+"DP", elig_type
  DropListBox 135, 5, 30, 15, "L"+chr(9)+"S"+chr(9)+"B", budget_type
  EditBox 285, 5, 85, 15, recipient_amt
  EditBox 90, 25, 280, 15, income
  EditBox 50, 45, 320, 15, deductions
  CheckBox 5, 65, 70, 10, "Updated RSPD?", updated_RSPD_check
  CheckBox 75, 65, 110, 10, "Approved new MAXIS results?", approved_check
  CheckBox 190, 65, 70, 10, "Sent DHS-3050?", DHS_3050_check
  CheckBox 5, 80, 125, 15, "Sent DHS-5181 to Case Manager", sent_5181_check
  EditBox 75, 100, 140, 15, designated_provider
  EditBox 75, 120, 295, 15, other
  DropListBox 60, 145, 60, 15, "None"+chr(9)+"Paperless IR"+chr(9)+"HRF", special_header_droplist
  EditBox 190, 145, 70, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 265, 145, 50, 15
    CancelButton 320, 145, 50, 15
    PushButton 220, 100, 35, 10, "ELIG/HC", ELIG_HC_button
    PushButton 265, 100, 25, 10, "BILS", BILS_button
    PushButton 290, 100, 25, 10, "FACI", FACI_button
    PushButton 315, 100, 25, 10, "HCMI", HCMI_button
    PushButton 340, 100, 25, 10, "UNEA", UNEA_button
  Text 5, 10, 35, 10, "Elig type:"
  Text 85, 10, 45, 10, "Budget type:"
  Text 175, 10, 110, 10, "Waiver obilgation/recipient amt:"
  Text 5, 30, 80, 10, "Total countable income:"
  Text 5, 50, 45, 10, "Deductions:"
  GroupBox 260, 90, 110, 25, "STAT based navigation"
  Text 5, 105, 70, 10, "Designated provider:"
  Text 5, 125, 65, 10, "Other (if applicable):"
  Text 5, 150, 55, 10, "Special header:"
  Text 130, 150, 60, 10, "Worker signature:"
EndDialog


'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connects to BlueZone
EMConnect ""

'Grabbing case number & footer month/year
Call MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

'Shows case number dialog
Dialog case_number_dialog
cancel_confirmation

'Sends transmit to check for MAXIS
Call check_for_MAXIS(FALSE)

'Going to ELIG/HC for the correct footer month
back_to_self
EMWriteScreen "elig", 16, 43
EMWriteScreen "________", 18, 43
EMWriteScreen MAXIS_case_number, 18, 43
EMWriteScreen MAXIS_footer_month, 20, 43
EMWriteScreen MAXIS_footer_year, 20, 46
EMWriteScreen "hc", 21, 70
transmit

'Checks to make sure it's in HCMI, due to error prone cases
call navigate_to_MAXIS_screen("STAT", "HCMI")
EMReadScreen HCMI_check, 4, 2, 55
If HCMI_check <> "HCMI" then transmit

'Checks the spenddown option. If one is indicated it will navigate to FACI and pull the current FACI into the designated provider box. If no FACI is given it will generate a warning message to the worker to check MMIS.
EMReadScreen spenddown_option, 2, 10, 57
If spenddown_option <> "__" then
  call autofill_editbox_from_MAXIS(array("01"), "FACI", FACI)
  If FACI = "" then
    MsgBox "A current facility could not be found. Check MMIS for a designated provider."
  Else
    designated_provider = FACI
  End if
End if

'Jumps into ELIG/HC for the footer month listed earlier
Call navigate_to_MAXIS_screen("ELIG", "HC__")

'Checks to see if MEMB 01 has a case. If not it'll try MEMB 02. If that doesn't work the script will error out on its own due to MAXIS intervention.
EMReadScreen person_check, 2, 8, 31
If person_check = "NO" then
  MsgBox "Person 01 does not have HC on this case. The script will attempt to execute this on person 02. Please check this for errors before approving any results."
  EMWriteScreen "x", 9, 26
End if
If person_check <> "NO" then EMWriteScreen "x", 8, 26
transmit

'Searching for the footer month span after defining row/col variables. If a span can't be found the script will shut down.
row = 1
col = 1
EMSearch " " & MAXIS_footer_month & "/" & MAXIS_footer_year & " ", row, col
If row = 0 then script_end_procedure("A " & MAXIS_footer_month & "/" & MAXIS_footer_year & " span could not be found. Try this again. You may need to run the case through background.")

'Grabbing the elig type and budget type
EMReadScreen elig_type, 2, 12, col - 1
EMReadScreen budget_type, 1, 13, col + 3
If (budget_type <> "L" AND budget_type <> "S" AND budget_type <> "B") THEN
	script_end_procedure ("This case is not a L, S or B budget case.  Use the ""Approved Programs"" script instead.")
END if

'Transmitting into the budget breakdown screen
EMWriteScreen "x", 9, col + 3
transmit

'Checking to see if this is an LBUD. If so, it'll grab the info from the appropriate places.
EMReadScreen LBUD_check, 4, 3, 45
If LBUD_check = "LBUD" then
  EMReadScreen recipient_amt, 10, 15, 70
  recipient_amt = "$" & trim(recipient_amt)
  EMReadScreen income, 10, 12, 32
  income = "$" & trim(income)
  EMReadScreen LTC_exclusions, 10, 14, 32
  If LTC_exclusions <> "__________" then deductions = deductions & "LTC exclusions ($" & replace(LTC_exclusions, "_", "") & "). "
  EMReadScreen medicare_premium, 10, 15, 32
  If medicare_premium <> "__________" then deductions = deductions & "Medicare ($" & replace(medicare_premium, "_", "") & "). "
  EMReadScreen pers_cloth_needs, 10, 16, 32
  If pers_cloth_needs <> "__________" then deductions = deductions & "Personal needs ($" & replace(pers_cloth_needs, "_", "") & "). "
  EMReadScreen home_maintenance_allowance, 10, 17, 32
  If home_maintenance_allowance <> "__________" then deductions = deductions & "Home maintenance allowance ($" & replace(home_maintenance_allowance, "_", "") & "). "
  EMReadScreen guard_rep_payee_fee, 10, 18, 32
  If guard_rep_payee_fee <> "__________" then deductions = deductions & "Payee fee ($" & replace(guard_rep_payee_fee, "_", "") & "). "
  EMReadScreen spousal_allocation, 10, 8, 70
  If spousal_allocation <> "          " then deductions = deductions & "Spousal allocation ($" & replace(spousal_allocation, " ", "") & "). "
  EMReadScreen family_allocation, 10, 9, 70
  If family_allocation <> "__________" then deductions = deductions & "Family allocation ($" & replace(family_allocation, "_", "") & "). "
  EMReadScreen health_ins_premium, 10, 10, 70
  If health_ins_premium <> "__________" then deductions = deductions & "Health insurance premium ($" & replace(health_ins_premium, "_", "") & "). "
  EMReadScreen other_med_expense, 10, 11, 70
  If other_med_expense <> "__________" then deductions = deductions & "Other medical expense ($" & replace(other_med_expense, "_", "") & "). "
  EMReadScreen SSI_1611_benefits, 10, 12, 70
  If SSI_1611_benefits <> "__________" then deductions = deductions & "SSI 1611 benefits ($" & replace(SSI_1611_benefits, "_", "") & "). "
  EMReadScreen other_deductions, 10, 13, 70
  If other_deductions <> "__________" then deductions = deductions & "Other deductions ($" & replace(other_deductions, "_", "") & "). "
End if

'Now it checks to see if this is an SBUD. If so, it'll grab the info from the appropriate places.
EMReadScreen SBUD_check, 4, 3, 44
If SBUD_check = "SBUD" then
  EMReadScreen recipient_amt, 10, 16, 71
  recipient_amt = "$" & trim(recipient_amt)
  EMReadScreen income, 10, 13, 32
  income = "$" & trim(income)
  EMReadScreen LTC_exclusions, 10, 15, 32
  If LTC_exclusions <> "__________" then deductions = deductions & "LTC exclusions ($" & replace(LTC_exclusions, "_", "") & "). "
  EMReadScreen medicare_premium, 10, 16, 32
  If medicare_premium <> "__________" then deductions = deductions & "Medicare ($" & replace(medicare_premium, "_", "") & "). "
  EMReadScreen pers_cloth_needs, 10, 17, 32
  If pers_cloth_needs <> "__________" then deductions = deductions & "Maintenance needs allowance ($" & replace(pers_cloth_needs, "_", "") & "). "
  EMReadScreen guard_rep_payee_fee, 10, 18, 32
  If guard_rep_payee_fee <> "__________" then deductions = deductions & "Payee fee ($" & replace(guard_rep_payee_fee, "_", "") & "). "
  EMReadScreen spousal_allocation, 10, 9, 71
  If spousal_allocation <> "          " then deductions = deductions & "Spousal allocation ($" & replace(spousal_allocation, " ", "") & "). "
  EMReadScreen family_allocation, 10, 10, 71
  If family_allocation <> "__________" then deductions = deductions & "Family allocation ($" & replace(family_allocation, "_", "") & "). "
  EMReadScreen health_ins_premium, 10, 11, 71
  If health_ins_premium <> "__________" then deductions = deductions & "Health insurance premium ($" & replace(health_ins_premium, "_", "") & "). "
  EMReadScreen other_med_expense, 10, 12, 71
  If other_med_expense <> "__________" then deductions = deductions & "Other medical expense ($" & replace(other_med_expense, "_", "") & "). "
  EMReadScreen SSI_1611_benefits, 10, 13, 71
  If SSI_1611_benefits <> "__________" then deductions = deductions & "SSI 1611 benefits ($" & replace(SSI_1611_benefits, "_", "") & "). "
  EMReadScreen other_deductions, 10, 14, 71
  If other_deductions <> "__________" then deductions = deductions & "Other deductions ($" & replace(other_deductions, "_", "") & "). "
End if

'Now it checks to see if this is an EBUD. If so, it'll grab the info from the appropriate places.
EMReadScreen EBUD_check, 4, 3, 60
If EBUD_check = "EBUD" then
  EMReadScreen income, 10, 9, 69
  income = "$" & trim(income)
  EMReadScreen MA_EPD_premium, 10, 13, 69
  other = "MA-EPD premium is $" & trim(MA_EPD_premium) & "/mo."
End if

'Now it checks to see if this is a BBUD. If so, it'll read the info, then offer the worker the chance to navigate to BILS
EMReadScreen BBUD_check, 4, 3, 47
If BBUD_check = "BBUD" then
  EMReadScreen income, 10, 12, 32
  income = "$" & trim(income)
  Dialog BBUD_dialog
  cancel_confirmation
  If ButtonPressed = 4 then
    PF3
    Call check_for_MAXIS(False)
      Do
        Dialog BBUD_Dialog
        cancel_confirmation
      Loop until MAXIS_check = "MAXIS"
    End if
    Call navigate_to_MAXIS_screen("STAT", "BILS")
    EMReadScreen BILS_check, 4, 2, 54
    If BILS_check <> "BILS" then transmit
End if

'auto-fills and cleans up information to entered into the approval_dialog
If recipient_amt = "$" or recipient_amt = "" then recipient_amt = "$0"
If income = "$" or income = "" then income = "$0"
If deductions = "$" or deductions = "" then deductions = "$0"


'Shows the MA approval dialog, checks for MAXIS and allows navigation buttons
Do
	err_msg = ""
	Dialog approval_dialog
	cancel_confirmation
	MAXIS_dialog_navigation
	If worker_signature = "" then err_msg = err_msg & vbNewLine & "* Sign your case note."
	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
LOOP until err_msg = ""


'checking for an active MAXIS session
Call check_for_MAXIS (FALSE)

'THE CASE NOTE----------------------------------------------------------------------------------------------------
start_a_blank_CASE_NOTE
'if case is L budget
If (special_header_droplist = "HRF" AND budget_type = "L") then Call write_variable_in_CASE_NOTE("**Approved " & elig_type & "-" & budget_type & " for HRF " & MAXIS_footer_month & "/" & MAXIS_footer_year & ", " & recipient_amt & " LTC SD**")
If (special_header_droplist = "Paperless IR" AND budget_type = "L") then Call write_variable_in_CASE_NOTE("**Approved " & elig_type & "-" & budget_type & " for paperless IR for " & MAXIS_footer_month & "/" & MAXIS_footer_year & ", " & recipient_amt & " LTC SD**")
If (special_header_droplist = "None" AND budget_type = "L") then Call write_variable_in_case_note("**Approved " & elig_type & "-" & budget_type & " for " & MAXIS_footer_month & "/" & MAXIS_footer_year & ", " & recipient_amt & " LTC SD**")
'if case is S budget
If (special_header_droplist = "HRF" AND budget_type = "S") then Call write_variable_in_CASE_NOTE("**Approved HRF " & elig_type & "-" & budget_type & " for HRF " & MAXIS_footer_month & "/" & MAXIS_footer_year & ", " & recipient_amt & " SISEW waiver obl**")
If (special_header_droplist = "Paperless IR" AND budget_type = "S") then Call write_variable_in_CASE_NOTE("**Approved " & elig_type & "-" & budget_type & " for paperless IR for " & MAXIS_footer_month & "/" & MAXIS_footer_year & ", " & recipient_amt & " SISEW waiver obl**")
If (special_header_droplist = "None" AND budget_type = "S") then Call write_variable_in_case_note("**Approved " & elig_type & "-" & budget_type & " for " & MAXIS_footer_month & "/" & MAXIS_footer_year & ", " & recipient_amt & " SISEW waiver obl**")
'if case is B budget
If (special_header_droplist = "HRF" AND budget_type = "B") then Call write_variable_in_CASE_NOTE("**Approved HRF " & elig_type & "-" & budget_type & " for HRF " & MAXIS_footer_month & "/" & MAXIS_footer_year & ", " & recipient_amt & " recip amt**")
If (special_header_droplist = "Paperless IR" AND budget_type = "B") then Call write_variable_in_CASE_NOTE("**Approved " & elig_type & "-" & budget_type & " for paperless IR for " & MAXIS_footer_month & "/" & MAXIS_footer_year & ", " & recipient_amt & " recip amt**")
If (special_header_droplist = "None" AND budget_type = "B") then Call write_variable_in_case_note("**Approved " & elig_type & "-" & budget_type & " for " & MAXIS_footer_month & "/" & MAXIS_footer_year & ", " & recipient_amt & " recip amt**")
call write_bullet_and_variable_in_case_note ("Income", income)
call write_bullet_and_variable_in_case_note ("Deductions", deductions)
call write_variable_in_case_note ("")
If updated_RSPD_check = 1 then call write_variable_in_case_note("* Updated RSPD in MMIS.")
call write_bullet_and_variable_in_case_note ("Designated provider", designated_provider)
If approved_check = 1 then call write_variable_in_case_note ("* Approved new MAXIS results.")
If DHS_3050_check = 1 then call write_variable_in_case_note ("* Sent DHS-3050 LTC communication form to facility.")
IF sent_5181_check = 1 then call write_variable_in_case_note ("* Sent DHS-5181 LTC communication to Case Manager")
call write_bullet_and_variable_in_case_note ("Other", other)
call write_variable_in_case_note ("---")
call write_variable_in_case_note (worker_signature)

script_end_procedure("")
