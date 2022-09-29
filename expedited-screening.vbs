'Required for statistical purposes==========================================================================================
'name_of_script = "NOTES - EXPEDITED SCREENING.vbs"
'start_time = timer
'STATS_counter = 1               'sets the stats counter at one
'STATS_manualtime = 180          'manual run time in seconds
'STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

'Because we are running these locally, we are going to get rid of all the calls to GitHub...
if (func_lib_run <> true OR IsEmpty(FuncLib_URL) = TRUE) then 
	FuncLib_URL = "I:\Blue Zone Scripts\Functions Library.vbs"
	Set run_funclib = CreateObject("Scripting.FileSystemObject")
	Set fso_funclib_command = run_funclib.OpenTextFile(FuncLib_URL)
	text_from_the_other_script = fso_funclib_command.ReadAll
	fso_funclib_command.Close
	Execute text_from_the_other_script
	func_lib_run = true
end if
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
'changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
'call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
'changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog exp_screening_dialog, 0, 0, 181, 210, "Expedited Screening Dialog"
  EditBox 55, 5, 95, 15, MAXIS_case_number
  EditBox 100, 25, 50, 15, income
  EditBox 100, 45, 50, 15, assets
  EditBox 100, 65, 50, 15, rent
  CheckBox 15, 95, 55, 10, "Heat (or AC)", heat_AC_check
  CheckBox 75, 95, 45, 10, "Electricity", electric_check
  CheckBox 130, 95, 35, 10, "Phone", phone_check
  DropListBox 70, 115, 105, 15, "intake"+chr(9)+"add-a-program", application_type
  EditBox 70, 135, 105, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 70, 155, 50, 15
    CancelButton 125, 155, 50, 15
  Text 10, 180, 160, 15, "The income, assets and shelter costs fields will default to $0 if left blank. "
  Text 5, 30, 95, 10, "Income received this month:"
  Text 5, 50, 95, 10, "Cash, checking, or savings: "
  Text 5, 70, 90, 10, "AMT paid for rent/mortgage:"
  GroupBox 5, 85, 170, 25, "Utilities claimed (check below):"
  Text 5, 120, 60, 10, "Application is for:"
  Text 5, 140, 60, 10, "Worker signature:"
  Text 5, 10, 50, 10, "Case number: "
  GroupBox 0, 170, 175, 30, "**IMPORTANT**"
EndDialog

'DATE BASED LOGIC FOR UTILITY AMOUNTS------------------------------------------------------------------------------------------
If application_date >= cdate("10/01/2020") then     'these variables need to change every October per CM.18.15.09
    heat_AC_amt = 496
    electric_amt = 154
    phone_amt = 56
ElseIf application_date >= cdate("10/01/2019") then
    'October 2019 amounts 
    heat_AC_amt = 490
    electric_amt = 143
    phone_amt = 49
elseIf date >= cdate("10/01/2018") then
	heat_AC_amt = 493
	electric_amt = 126
	phone_amt = 47
ElseIf date >= cdate("10/01/2017") then 
	heat_AC_amt = 536
	electric_amt = 172
	phone_amt = 41
ElseIf date >= cdate("10/01/2016") then			'these variables need to change every October
	heat_AC_amt = 532
	electric_amt = 141
	phone_amt = 38
Else
	heat_AC_amt = 454
	electric_amt = 141
	phone_amt = 38
End if

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connecting to BlueZone
EMConnect ""
'It will search for a case number.
call MAXIS_case_number_finder(MAXIS_case_number)

dim confirm_cancel

'Shows the dialog
Do
	Do
		Do
			Dialog exp_screening_dialog
			if ButtonPressed = 0 then 
				confirm_cancel = MsgBox ("Are you sure you want to cancel? Press YES to return to the previous script. Press NO to continue.", vbYesNo)
				if confirm_cancel = vbYes THEN exit do
			end if
			If isnumeric(MAXIS_case_number) = False then MsgBox "You must enter a valid case number."
		Loop until isnumeric(MAXIS_case_number) = True
		if confirm_cancel = vbYes THEN exit do
		If (income <> "" and isnumeric(income) = false) or (assets <> "" and isnumeric(assets) = false) or (rent <> "" and isnumeric(rent) = false) then MsgBox "The income/assets/rent fields must be numeric only. Do not put letters or symbols in these sections."
	Loop until (income = "" or isnumeric(income) = True) and (assets = "" or isnumeric(assets) = True) and(rent = "" or isnumeric(rent) = True)
	if confirm_cancel = vbYes then exit do
	If worker_signature = "" then MsgBox "You must sign your case note."
Loop until worker_signature <> ""

if confirm_cancel <> vbYes THEN 

	'checking for an active MAXIS session
	Call check_for_MAXIS(FALSE)

'LOGIC AND CALCULATIONS----------------------------------------------------------------------------------------------------
'Logic for figuring out utils. The highest priority for the if...then is heat/AC, followed by electric and phone, followed by phone and electric separately.
If heat_AC_check = checked then
	utilities = heat_AC_amt
ElseIf electric_check = checked and phone_check = checked then
	utilities = phone_amt + electric_amt					'Phone standard plus electric standard.
ElseIf phone_check = checked and electric_check = unchecked then
	utilities = phone_amt
ElseIf electric_check = checked and phone_check = unchecked then
	utilities = electric_amt
End if

'in case no options are clicked, utilities are set to zero.
If phone_check = unchecked and electric_check = unchecked and heat_AC_check = unchecked then utilities = 0

'If nothing is written for income/assets/rent info, we set to zero.
If trim(income) = "" then income = 0
If trim(assets) = "" then assets = 0
If trim(rent) = "" then rent = 0

'Calculates expedited status based on above numbers
If (int(income) < 150 and int(assets) <= 100) or ((int(income) + int(assets)) < (int(rent) + cint(utilities))) then expedited_status = "client appears expedited"
If (int(income) + int(assets) >= int(rent) + cint(utilities)) and (int(income) >= 150 or int(assets) > 100) then expedited_status = "client does not appear expedited"
'----------------------------------------------------------------------------------------------------

'Navigates to STAT/DISQ using current month as footer month. If it can't get in to the current month due to CAF received in a different month, it'll find that month and navigate to it.
Call navigate_to_MAXIS_screen("STAT", "DISQ")
'grabbing footer month and year
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

'Reads the DISQ info for the case note.
EMReadScreen DISQ_member_check, 34, 24, 2
If DISQ_member_check = "DISQ DOES NOT EXIST FOR ANY MEMBER" then
	has_DISQ = False
Else
	has_DISQ = True
End if

'Reads MONY/DISB to see if EBT account is open
IF expedited_status = "client appears expedited" THEN
	Call navigate_to_MAXIS_screen("MONY", "DISB")
	EMReadScreen EBT_account_status, 1, 14, 27
END IF

'THE CASE NOTE----------------------------------------------------------------------------------------------------
	call navigate_to_MAXIS_screen("case", "note")
	PF9

	EMReadScreen case_note_check, 17, 2, 33
	EMReadScreen mode_check, 1, 20, 09
	If case_note_check <> "Case Notes (NOTE)" or mode_check <> "A" then    'this will account for those cases when the script is run on an out of county case.
		msgbox "The script can't open a case note. You may be in inquiry or entered a case number that is in another county." &_
		vbNewLine & vbNewLine & "This result for this case is " & expedited_status & vbNewLine & vbNewLine & "Please run the script again if you were in inquiry to add a case note."
		script_end_procedure("")
	else
		'Body of the case note
		Call write_variable_in_CASE_NOTE("Received " & application_type & ", " & expedited_status)
		call write_variable_in_CASE_NOTE("---")
		call write_variable_in_CASE_NOTE("     CAF 1 income claimed this month: $" & income)
		call write_variable_in_CASE_NOTE("         CAF 1 liquid assets claimed: $" & assets)
		call write_variable_in_CASE_NOTE("         CAF 1 rent/mortgage claimed: $" & rent)
		call write_variable_in_CASE_NOTE("        Utilities (amt/HEST claimed): $" & utilities)
		call write_variable_in_CASE_NOTE("---")
		If has_DISQ = True then call write_variable_in_CASE_NOTE("A DISQ panel exists for someone on this case.")
		If has_DISQ = False then call write_variable_in_CASE_NOTE("No DISQ panels were found for this case.")
		If expedited_status = "client appears expedited" AND EBT_account_status = "Y" then call write_variable_in_CASE_NOTE("* EBT Account IS open.  Recipient will NOT be able to get a replacement card in the agency.  Rapid Electronic Issuance (REI) with caution.")
		If expedited_status = "client appears expedited" AND EBT_account_status = "N" then call write_variable_in_CASE_NOTE("* EBT Account is NOT open.  Recipient is able to get initial card in the agency.  Rapid Electronic Issuance (REI) can be used, but only to avoid an emergency issuance or to meet EXP criteria.")
		call write_variable_in_CASE_NOTE("---")
		call write_variable_in_CASE_NOTE(worker_signature)
		If expedited_status = "client appears expedited" then
			MsgBox "This client appears expedited. A same day interview needs to be offered."
		End if
		If expedited_status = "client does not appear expedited" then
			MsgBox "This client does not appear expedited. A same day interview does not need to be offered."
		End if
	End if

end if
'script_end_procedure("")
