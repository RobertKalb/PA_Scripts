'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "UTILITIES - MAIN MENU.vbs"
start_time = timer

'Because we are running these locally, we are going to get rid of all the calls to GitHub...
' FuncLib_URL = "I:\Blue Zone Scripts\Functions Library.vbs"
' Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
' Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
' text_from_the_other_script = fso_command.ReadAll
' fso_command.Close
' Execute text_from_the_other_script
' func_lib_run = true
' 'END FUNCTIONS LIBRARY BLOCK================================================================================================

FUNCTION launch_selected_script(script_file_path)
  Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
  Set fso_command = run_another_script_fso.OpenTextFile(script_file_path)
  text_from_the_other_script = fso_command.ReadAll
  fso_command.Close
  ExecuteGlobal text_from_the_other_script
END FUNCTION

Set objNet = CreateObject("WScript.NetWork") 
windows_user_ID = UCASE(objNet.UserName)

'A class for each script item
class script

	public script_name             	'The familiar name of the script
	public file_name               	'The actual file name
	public description             	'The description of the script
	public button                  	'A variable to store the actual results of ButtonPressed (used by much of the script functionality)
    public category               	'The script category (ACTIONS/BULK/etc)
    public SIR_instructions_URL    	'The instructions URL in SIR
    public button_plus_increment	'Workflow scripts use a special increment for buttons (adding or subtracting from total times to run). This is the add button.
	public button_minus_increment	'Workflow scripts use a special increment for buttons (adding or subtracting from total times to run). This is the minus button.
	public total_times_to_run		'A variable for the total times the script should run
	public subcategory				'An array of all subcategories a script might exist in, such as "LTC" or "A-F"

	public property get button_size	'This part determines the size of the button dynamically by determining the length of the script name, multiplying that by 3.5, rounding the decimal off, and adding 10 px
		button_size = round ( len( script_name ) * 3.5 ) + 10
	end property

end class

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
'changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
'call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
'changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'Because we are running these locally, we are going to get rid of all the calls to GitHub...
script_list_URL = "I:\Blue Zone Scripts\Public Assistance Script Files\DHS-MAXIS-Scripts-master\COMPLETE LIST OF SCRIPTS.vbs"
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile(script_list_URL)
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script


Function declare_main_menu_dialog(script_category, ButtonPressed)

	'Runs through each script in the array and generates a list of subcategories based on the category located in the function. Also modifies the script description if it's from the last two months, to include a "NEW!!!" notification.
	For current_script = 0 to ubound(script_array)
		'Subcategory handling (creating a second list as a string which gets converted later to an array)
		If ucase(script_array(current_script).category) = ucase(script_category) then																								'If the script in the array is of the correct category (ACTIONS/NOTES/ETC)...
			For each listed_subcategory in script_array(current_script).subcategory																									'...then iterate through each listed subcategory, and...
				If listed_subcategory <> "" and InStr(subcategory_list, ucase(listed_subcategory)) = 0 then subcategory_list = subcategory_list & "|" & ucase(listed_subcategory)	'...if the listed subcategory isn't blank and isn't already in the list, then add it to our handy-dandy list.
			Next
		End if
		'Adds a "NEW!!!" notification to the description if the script is from the last two months.
		If DateDiff("m", script_array(current_script).release_date, DateAdd("m", -2, date)) <= 0 then
			script_array(current_script).description = "NEW " & script_array(current_script).release_date & "!!! --- " & script_array(current_script).description
			script_array(current_script).release_date = "12/12/1999" 'backs this out and makes it really old so it doesn't repeat each time the dialog loops. This prevents NEW!!!... from showing multiple times in the description.
		End if

	Next

	subcategory_list = split(subcategory_list, "|")
	total_number_of_subcategories = ubound(subcategory_list)
	ReDim subcategory_array(total_number_of_subcategories, 1)

	For i = 0 to ubound(subcategory_list)

		'set subcategory_array(i) = new subcat
		If subcategory_list(i) = "" then subcategory_list(i) = "MAIN"
		subcategory_array(i, 0) = subcategory_list(i)
	Next

	BeginDialog dialog1, 0, 0, 600, 400, script_category & " scripts main menu dialog"
	 	Text 5, 5, 435, 10, script_category & " scripts main menu: select the script to run from the choices below."
	  	ButtonGroup ButtonPressed

		'SUBCATEGORY HANDLING--------------------------------------------
		subcat_button_position = 5

		For i = 0 to ubound(subcategory_array)

			'Displays the button and text description-----------------------------------------------------------------------------------------------------------------------------
			'FUNCTION		HORIZ. ITEM POSITION	VERT. ITEM POSITION		ITEM WIDTH	ITEM HEIGHT		ITEM TEXT/LABEL				BUTTON VARIABLE
			PushButton 		subcat_button_position, 20, 					50, 		15, 			subcategory_array(i, 0), 	subcat_button_placeholder

			subcategory_array(i, 1) = subcat_button_placeholder	'The .button property won't carry through the function. This allows it to escape the function. Thanks VBScript.
			subcat_button_position = subcat_button_position + 50
			subcat_button_placeholder = subcat_button_placeholder + 1
		Next


		'SCRIPT LIST HANDLING--------------------------------------------


		'' 	PushButton 445, 10, 65, 10, "SIR instructions", 	SIR_instructions_button
		'This starts here, but it shouldn't end here :)
		vert_button_position = 50


		For current_script = 0 to ubound(script_array)
			If ucase(script_array(current_script).category) = ucase(script_category) then

				'Joins all subcategories together
				subcategory_string = ucase(join(script_array(current_script).subcategory))

				'Accounts for scripts without subcategories
				If subcategory_string = "" then subcategory_string = "MAIN"		'<<<THIS COULD BE A PROPERTY OF THE CLASS

				'If the selected subcategory is in the subcategory string, it will display those scripts
				If InStr(subcategory_string, subcategory_selected) <> 0 then

					SIR_button_placeholder = button_placeholder + 1	'We always want this to be one more than the button_placeholder

					'Displays the button and text description-----------------------------------------------------------------------------------------------------------------------------
					'FUNCTION		HORIZ. ITEM POSITION	VERT. ITEM POSITION		ITEM WIDTH	ITEM HEIGHT		ITEM TEXT/LABEL										BUTTON VARIABLE
					PushButton 		5, 						vert_button_position, 	10, 		10, 			"?", 												SIR_button_placeholder
					PushButton 		18,						vert_button_position, 	120, 		10, 			script_array(current_script).script_name, 			button_placeholder
					Text 			120 + 23, 				vert_button_position, 	500, 		10, 			"--- " & script_array(current_script).description
					'----------
					vert_button_position = vert_button_position + 15	'Needs to increment the vert_button_position by 15px (used by both the text and buttons)
					'----------
					script_array(current_script).button = button_placeholder	'The .button property won't carry through the function. This allows it to escape the function. Thanks VBScript.
					script_array(current_script).SIR_instructions_button = SIR_button_placeholder	'The .button property won't carry through the function. This allows it to escape the function. Thanks VBScript.
					button_placeholder = button_placeholder + 2
				End if
			End if
		next
		CancelButton 540, 380, 50, 15
	EndDialog
End function

'Starting these with a very high number, higher than the normal possible amount of buttons.
'	We're doing this because we want to assign a value to each button pressed, and we want
'	that value to change with each button. The button_placeholder will be placed in the .button
'	property for each script item. This allows it to both escape the Function and resize
'	near infinitely. We use dummy numbers for the other selector buttons for much the same reason,
'	to force the value of ButtonPressed to hold in near infinite iterations.
button_placeholder 			= 24601
subcat_button_placeholder 	= 1701

'Other pre-loop and pre-function declarations
Dim subcategory_array : subcategory_array = Array()
subcategory_string = ""
subcategory_selected = "MAIN"

'Defining dialog1 as 1000. Assigning a numeric value seems to work to preserve a high amount of buttons for our scripts.
dialog1 = 1000

'Displays the dialog
Do

	'Creates the dialog
	call declare_main_menu_dialog("Utilities", ButtonPressed)

	'At the beginning of the loop, we are not ready to exit it. Conditions later on will impact this.
	ready_to_exit_loop = false

	'Displays dialog, if cancel is pressed then stopscript
	dialog
	If ButtonPressed = 0 then stopscript

	'Determines the subcategory if a subcategory button was selected.
	For i = 0 to ubound(subcategory_array)
		If ButtonPressed = subcategory_array(i, 1) then subcategory_selected = subcategory_array(i, 0)
	Next

	'Runs through each script in the array... if the user selected script instructions (via ButtonPressed) it'll open_URL_in_browser to those instructions
	For i = 0 to ubound(script_array)
		If ButtonPressed = script_array(i).SIR_instructions_button then call open_URL_in_browser(script_array(i).SIR_instructions_URL)
	Next

	'Runs through each script in the array... if the user selected the actual script (via ButtonPressed), it'll run_from_GitHub
	For i = 0 to ubound(script_array)
		If ButtonPressed = script_array(i).button then
			ready_to_exit_loop = true		'Doing this just in case a stopscript or script_end_procedure is missing from the script in question
			script_to_run = script_array(i).script_URL
			Exit for
		End if
	Next


Loop until ready_to_exit_loop = true

'Updating dialog1 to be a separate numeric value. This might not be necessary but it's currently working so I am not changing it.
dialog1 = dialog1 + 1

call launch_selected_script(script_to_run)

stopscript
