'STATS GATHERING--------------------------------------------------------------------------------------------------------------
name_of_script = "UTILITIES - UPDATE WORKER SIGNATURE.vbs"
start_time = timer

' 'Because we are running these locally, we are going to get rid of all the calls to GitHub...
' FuncLib_URL = "I:\Blue Zone Scripts\Functions Library.vbs"
' Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
' Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
' text_from_the_other_script = fso_command.ReadAll
' fso_command.Close
' Execute text_from_the_other_script
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
' 'END CHANGELOG BLOCK =======================================================================================================

'Needs to determine MyDocs directory before proceeding.
Set wshshell = CreateObject("WScript.Shell")
user_myDocs_folder = wshShell.SpecialFolders("MyDocuments") & "\"

'Looks for the file. If found, it alerts the user
Dim oTxtFile
With (CreateObject("Scripting.FileSystemObject"))
	If .FileExists(user_myDocs_folder & "workersig.txt") Then
		Set get_worker_sig = CreateObject("Scripting.FileSystemObject")
		Set worker_sig_command = get_worker_sig.OpenTextFile(user_myDocs_folder & "workersig.txt")
		worker_sig = worker_sig_command.ReadAll
		IF worker_sig <> "" THEN worker_signature = worker_sig
		worker_sig_command.Close

		worker_signature_msgbox = MsgBox("A worker signature was found! You are listed as: " & worker_signature & "." & vbNewLine & vbNewLine & _
			"Your worker signature was found at: " & user_myDocs_folder & "workersig.txt." & vbNewLine & vbNewLine & _
			"Would you like to update this signature? Press Yes to continue, or No to cancel.", vbYesNo + vbQuestion, "A worker signature was found!")

		If worker_signature_msgbox = vbNo then StopScript
	END IF
END WITH

'----------DIALOGS----------
BeginDialog dialog1, 0, 0, 191, 105, "Update Worker Signature"
  EditBox 10, 60, 175, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 45, 85, 50, 15
    CancelButton 95, 85, 50, 15
  Text 10, 10, 175, 10, "Enter what you would like for your default signature."
  Text 10, 25, 170, 25, "NOTE: This will be pre-loaded in every script. Once the script has started, you can still modify your signature in the appropriate editbox."
EndDialog

'----------THE SCRIPT----------
dialog dialog1										'Shows the dialog
IF ButtonPressed = cancel THEN stopscript			'Handling for if cancel is pressed
IF worker_signature = "" THEN stopscript			'If they enter nothing, it exits

'This creates an object which collects the username from the Windows logon. We need this to determine the correct location for the My Documents folder.
Set objNet = CreateObject("WScript.NetWork")
windows_user_ID = objNet.UserName		'Saves the .UserName object as a new variable, windows_user_ID

'Opens an FSO, opens workersig.txt, writes the new signature in, and exits
SET update_worker_sig_fso = CreateObject("Scripting.FileSystemObject")
SET update_worker_sig_command = update_worker_sig_fso.CreateTextFile(user_myDocs_folder & "workersig.txt", 2)
update_worker_sig_command.Write(worker_signature)
update_worker_sig_command.Close

'Script ends
script_end_procedure("")
