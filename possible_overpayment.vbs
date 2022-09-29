'LOADING GLOBAL CLASSES--------------------------------------------------------------------
set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
set fso_command = run_another_script_fso.OpenTextFile("I:\Blue Zone Scripts\Script Files\global classes.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

MSGBOX "Script void."

' name_of_script = "ANOKA - DAIL - POSSIBLE OVERPAYMENT.vbs"
' start_time = timer
' 
' BeginDialog Dialog1, 0, 0, 246, 170, "Dialog"
'   ButtonGroup ButtonPressed
'     OkButton 140, 150, 50, 15
'     CancelButton 190, 150, 50, 15
'   Text 10, 15, 70, 10, "MAXIS Case Number"
'   EditBox 90, 10, 65, 15, maxis_case_number
'   Text 165, 15, 45, 10, "MEMB Num"
'   EditBox 215, 10, 25, 15, Edit2
' EndDialog
' 
' 
' 
' 
' set err_msg = new error_message
' do
' 	err_msg.reset_message
' 	Dialog
' 		cancel_confirmation
' 		if maxis_case_number 