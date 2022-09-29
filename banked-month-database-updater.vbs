'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "UTILITIES - BANKED MONTH DATABASE UPDATER.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 43               	'manual run time in seconds
STATS_denomination = "C"       		'C is for Case
'END OF stats block=========================================================================================================

'Because we are running these locally, we are going to get rid of all the calls to GitHub...
FuncLib_URL = "I:\Blue Zone Scripts\Functions Library.vbs"
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script
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


'Connects to BlueZone
EMConnect ""
'Making sure the county has the needed database, otherwise stop.
IF banked_months_db_tracking <> true THEN script_end_procedure("Your county must be using the MS-ACCESS ABAWD Banked month database to use this script.  The script will now stop.")

'THE SCRIPT-------------------------------------------------------------------------
	'Settng constants
		Const adSchemaColumns = 4
		'Creating objects for Access
		Set objConnection = CreateObject("ADODB.Connection")
		Set objRecordSet = CreateObject("ADODB.Recordset")


		'Opening DB
	objConnection.Open "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " & banked_month_database_path
	'Creating a recorset and collecting field names
 		Set objRecordSet = objConnection.OpenSchema(adSchemaColumns, Array(Null, Null, "banked_month_log"))

		Do until objRecordSet.EOF 'loop through all columns in recordset
			IF objRecordSet("Column_Name") <> "ID" AND objRecordSet("Column_Name") <> "MAXIS_case_number" AND objRecordSet("Column_Name") <> "member_number" then
				months_array = months_array & "," & cint(objRecordSet("Column_Name"))
			END If
			objRecordSet.MoveNext
		Loop
		months_array = split(months_array, ",")

	objConnection.Close
 	set objRecordSet = nothing

	call convert_array_to_droplist_items (months_array, months_list) 'This converts the array of months into a droplist for dialog

'dialogs
BeginDialog dialog1, 0, 0, 191, 105, "ABAWD BANKED MONTH DATABASE UPDATE"
  ButtonGroup ButtonPressed
    OkButton 75, 85, 50, 15
    CancelButton 130, 85, 50, 15
  DropListBox 125, 15, 45, 20, months_list, db_month
  Text 15, 10, 100, 20, "Choose the database month to evaluate:"
	Text 10, 35, 170, 35, "NOTE: This utility will check MAXIS case status and update the database.  It will only update the selected month.  For best results, run this utility after the desired month has ended."

EndDialog

'Connects to BlueZone
EMConnect ""

'Shows dialog
Dialog
If buttonpressed = cancel then stopscript

'setting footer month and year for MAXIS'
MAXIS_footer_year = "16"
MAXIS_footer_month = db_month
if len(MAXIS_footer_month) = 1 then MAXIS_footer_month = "0" & MAXIS_footer_month

'Checking for MAXIS
call check_for_maxis(false)

	'Connecting to the database
	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.Recordset")

	objConnection.Open "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " & banked_month_database_path & ""

	'creating a recordset of all active cases for selected month
		set rs = objConnection.Execute("SELECT * FROM banked_month_log WHERE " & db_month & " <> 0")
		rs.MoveFirst
		IF NOT(rs.EOF) THEN

			DO 'THis loop will look at ELIG to determine if this person was closed or remains open.
				IF rs("1") = true THEN
				STATS_counter = STATS_counter + 1 'add 1 to the stats count for each case checked
				MAXIS_case_number = rs("MAXIS_case_number") 'grab case number from current record
				member_number = rs("member_number") 'grab member_number'
				call navigate_to_MAXIS_screen("ELIG", "FS")
				' Make sure there is a version to read
				EMReadscreen version_exists, 10, 24, 2
				IF version_exists = "NO VERSION" THEN
					abawd_active = FALSE
				ELSE
				'Find most recent approved version
					EMReadScreen version, 2, 2, 18
					For approved = version to 0 Step -1
						EMReadScreen approved_check, 8, 3, 3
						If approved_check = "APPROVED" then Exit FOR
						version = version -1
						EMWriteScreen version, 19, 78
						transmit
						Next
						' Check to make sure that the member in question was eligible on most recent approval'
						IF len(member_number) = 1 THEN member_number = "0" & member_number
			  	abawd_active = true
					FOR i = 7 to 19 'this loop will look at each hh members elig factors'
						EMReadscreen ref_nbr, 2, i, 10
						IF ref_nbr = member_number THEN
							EMReadscreen member_test, 10, i, 57
							IF member_test = "INELIGIBLE" THEN abawd_active = false
							END If
							NEXT
					END IF
					'now go to WREG and check to make sure they are still coded ABAWD 10
					call navigate_to_MAXIS_screen("STAT", "WREG")
					EMWriteScreen member_number, 20, 76
					transmit
					EMReadScreen abawd_status, 2, 13, 50
					IF abawd_status <> "10" THEN abawd_active = false 'IF they aren't coded a 10, can't be a banked month, so clear this member from DB
				back_to_self

				'If not active, update the DB accordingly
				IF abawd_active = false THEN
				objConnection.Execute("UPDATE banked_month_log Set " & replace(MAXIS_footer_month, "0", "") & " = 0 WHERE MAXIS_case_number = " & MAXIS_case_number & " AND member_number = " & member_number &"")
				updated_case_list = updated_case_list & " " & MAXIS_case_number
				END IF
				END IF
			rs.MoveNext 'Switch to next record
			LOOP UNTIL(rs.eof = true)
		END IF
	objConnection.Close
	Set rs = nothing

STATS_counter = STATS_counter - 1                   'the count started at 1, should remove for accuracy
script_end_procedure("Success. The DB has been updated for inactive cases.  The following cases were updated: " & updated_case_list)
