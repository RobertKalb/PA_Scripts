'Required for statistical purposes==========================================================================================
name_of_script = "NOTICES - LTC - ASSET TRANSFER.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 70                               'manual run time in seconds
STATS_denomination = "C"       'C is for each CASE
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
' 
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

BeginDialog LTC_asset_transfer_dialog, 0, 0, 126, 82, "LTC asset transfer dialog"
  EditBox 35, 0, 85, 15, client
  EditBox 35, 20, 85, 15, spouse
  EditBox 70, 40, 50, 15, renewal_footer_month_year
  ButtonGroup LTC_asset_transfer_dialog_ButtonPressed
    OkButton 10, 60, 50, 15
    CancelButton 65, 60, 50, 15
  Text 5, 5, 30, 10, "Client:"
  Text 5, 25, 30, 10, "Spouse:"
  Text 5, 45, 65, 10, "ER date (MM/YY):"
EndDialog

'The script------------------------
'connecting to MAXIS
EMConnect ""
Do
  Dialog LTC_asset_transfer_dialog
  If LTC_asset_transfer_dialog_ButtonPressed = 0 then stopscript
  EMSendKey "<enter>"
  EMWaitReady 1, 1
  EMReadScreen WCOM_input_check, 27, 2, 28
  If WCOM_input_check <> "Worker Comment Input Screen" and WCOM_input_check <> "  Client Memo Input Screen " then MsgBox "You need to be on a notice in SPEC/WCOM or SPEC/MEMO for this to work. Please try again."
Loop until WCOM_input_check = "Worker Comment Input Screen" or WCOM_input_check = "  Client Memo Input Screen "

EMSendKey "<home>" + "The ownership of " + client + "'s assets must be transferred to " + spouse + " to avoid having them counted in future eligibility determinations. You are encouraged to do this as soon as possible. This transfer of assets must be done before " + client + "'s first annual renewal for " + renewal_footer_month_year + ". Verification of the transfer can be provided at any time. " + "<newline>" + "<newline>"
EMSendKey "At the first annual renewal in " + renewal_footer_month_year + " the value of all assets that list " + client + " as an owner or co-owner will be applied towards the Medical Assistance Asset limit of $3,000.00.  If the total value of all countable assets for " + client + " is more than $3,000.00, Medical Assistance may be closed for " + renewal_footer_month_year + "."

script_end_procedure("")
NOTICES
