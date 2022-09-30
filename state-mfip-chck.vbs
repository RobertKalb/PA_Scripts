'Required for statistical purposes===============================================================================
name_of_script = "state-mfip-chck.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 180                               'manual run time in seconds
STATS_denomination = "C"       'C is for each Case
'END OF stats block==============================================================================================

'LOADING FUNCTIONS LIBRARY FROM REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN  'Shouldn't load FuncLib if it already loaded once
                IF run_locally = FALSE or run_locally = "" THEN       'If the scripts are set to run locally, it skips this and uses an FSO below.
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
'("03/20/2020", "Experimental for COVID-19 pandemic.", "Robin Hoffman, MNIT with help from Jennifer Munger, DHS - ccm 43921")
'("10/16/2019", "All infrastructure changed to run locally and stored in BlueZone Scripts ccm. MNIT @ DHS)
'END CHANGELOG BLOCK =======================================================================================================

'Dialogs
'>>>>>Main dlg<<<<<
BeginDialog main_menu, 0, 0, 201, 70, "Benefit Month for MFIP/DW COVID-19 from List"
  EditBox 65, 25, 30, 15, MAXIS_footer_month
  EditBox 130, 25, 30, 15, MAXIS_footer_year
  ButtonGroup ButtonPressed
    OkButton 90, 40, 50, 15
    CancelButton 140, 40, 50, 15
   Text 10, 30, 50, 10, "Benefit month:"
  Text 100, 30, 25, 10, "Year:"
EndDialog

'>>>>> Function to build dlg for manual entry <<<<<
FUNCTION build_manual_entry_dlg(case_number_array, spec_memo, case_note_body, worker_signature)
                'Array for all case numbers
                'This was chosen over building a dlg with 50 variables
                REDim all_cases_array(50, 0)

                BeginDialog man_entry_dlg, 0, 0, 331, 330, "Enter MAXIS case numbers"
                                Text 10, 15, 140, 10, "Enter MAXIS case numbers below..."
                                dlg_row = 30
                                dlg_col = 10
                                FOR i = 1 TO 50
                                    EditBox dlg_col, dlg_row, 55, 15, all_cases_array(i, 0)
                                    dlg_row = dlg_row + 20
                                    IF dlg_row = 230 THEN
                                       dlg_row = 30
                                       dlg_col = dlg_col + 65
                                    END IF
                                NEXT
                                text 10, 235, 120, 10, "Enter case note below"
                                Text 10, 255, 25, 10, "MEMO:"
                                Text 10, 275, 20, 10, "Case Note:"
                                Text 10, 295, 60, 10, "Worker Signature:"
                                EditBox 45, 250, 280, 15, spec_memo
                                EditBox 35, 270, 290, 15, case_note_body
                                EditBox 75, 290, 150, 15, worker_signature
                                ButtonGroup ButtonPressed
                                                OkButton 220, 310, 50, 15
                                                CancelButton 270, 310, 50, 15
                EndDialog

                'Calling the dlg within the function
                DO
               	 'err_msg handling
                   err_msg = ""
                   DIALOG man_entry_dlg
                   cancel_confirmation
                   FOR i = 1 TO 50
                   	all_cases_array(i, 0) = replace(all_cases_array(i, 0), " ", "")
                        IF all_cases_array(i, 0) <> "" THEN
                           IF len(all_cases_array(i, 0)) > 8 THEN err_msg = err_msg & vbCr & "* Case number " & all_cases_array(i, 0) & " is too long to be a valid MAXIS case number."
                              IF isnumeric(all_cases_array(i, 0)) = FALSE THEN err_msg = err_msg & vbCr & "* Case number " & all_cases_array(i, 0) & " contains alphabetic characters. These are not valid."
                        END IF
                   NEXT
                   IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
                LOOP UNTIL err_msg = ""

                'building the array
                case_number_array = ""
                FOR i = 1 TO 50
                    IF all_cases_array(i, 0) <> "" THEN case_number_array = case_number_array & all_cases_array(i, 0) & "~~~"
                NEXT
END FUNCTION

'>>>>>DLG for Excel mode<<<<<
BeginDialog CASE_NOTE_from_excel_dlg, 0, 0, 256, 165, "MONY/CHCK Information"
  EditBox 220, 10, 25, 15, excel_col
  EditBox 65, 30, 40, 15, excel_row
  EditBox 190, 30, 40, 15, end_row
  EditBox 45, 50, 205, 15, case_note_header
  EditBox 35, 70, 215, 15, case_note_body
  EditBox 10, 98, 240, 15, memo_text
  EditBox 75, 120, 150, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 130, 145, 55, 15
    CancelButton 190, 145, 60, 15
  Text 10, 15, 205, 10, "Please enter the column containing the MAXIS case numbers..."
  Text 10, 35, 50, 10, "Row to start..."
  Text 135, 35, 50, 10, "Row to end..."
  Text 10, 55, 25, 10, "Header:"
  Text 10, 75, 20, 10, "Body:"
  Text 10, 125, 60, 10, "Worker Signature:"
  GroupBox 5, 90, 250, 30, "Special Memo - use semi-colon btw lines"
EndDialog

'----------FUNCTIONS----------
'-----This function needs to be added to the FUNCTIONS FILE-----
'>>>>> This function converts the letter for a number so the script can work with it <<<<<
FUNCTION convert_excel_letter_to_excel_number(excel_col)
                IF isnumeric(excel_col) = FALSE THEN
                                alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
                                excel_col = ucase(excel_col)
                                IF len(excel_col) = 1 THEN
                                   excel_col = InStr(alphabet, excel_col)
                                ELSEIF len(excel_col) = 2 THEN
                                   excel_col = (26 * InStr(alphabet, left(excel_col, 1))) + (InStr(alphabet, right(excel_col, 1)))
                                END IF
                ELSE
                                excel_col = CInt(excel_col)
                END IF
END FUNCTION

'----------MONY/CHCK----------
'-----These functions are setup to read the MONY/CHCK file for each type of program
'>>>>> This function  will be run from the call to it for CASH I active / MF or DW <<<<<
FUNCTION create_mony_chck_mf()
                                ' Count the members in the case
								DD_case = ""
                                                household_persons = ""
                                                client_array = ""
                                                test_array = ""
                                                pers_count = 0
                                                total_clients = 0
                                                'Create hh_member_array
                                                                CALL Navigate_to_MAXIS_screen("STAT", "MEMB")   'navigating to stat memb to gather the ref number and name.

                DO                                                                                                                          'reads the reference number, last name, first name, and then puts it into a single string then into the array
                                EMReadscreen ref_nbr, 3, 4, 33
                                EMReadscreen last_name, 5, 6, 30
                                EMReadscreen first_name, 7, 6, 63
                                EMReadscreen Mid_intial, 1, 6, 79
                                EMReadscreen age, 2, 8, 76
                                last_name = replace(last_name, "_", "") & " "
                                first_name = replace(first_name, "_", "") & " "
                                mid_initial = replace(mid_initial, "_", "")
                                client_string = ref_nbr & last_name & first_name & mid_intial 
                                client_array = client_array & client_string & "|"
                                age_string = 0
                                IF isnumeric(age) = FALSE THEN 
                                                age_string = 1
                                ELSE
                                                age_string = age
                                END IF
                                index_array = index_array & age_string & "|"
                                
                                transmit
                                Emreadscreen edit_check, 7, 24, 2
                LOOP until edit_check = "ENTER A"                                          'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.
                
                age_array = split(index_array, "|")
                test_array = split(client_array, "|")
                total_clients = Ubound(test_array)                                         'setting the upper bound for how many spaces to use from the array          
                
                'this is to search the excel spreadsheet for the matching case number and move the amount in the next column
                WHAT_TO_FIND = MAXIS_case_number                                                         
                Set FoundCell = objExcel.Range("A1:A550").Find(WHAT_TO_FIND)
                row=FoundCell.Row
                col=FoundCell.Column                              
                'MsgBox (MAXIS_case_number & ":  " & objExcel.Cells(row,col+1))
                mfip_amount =  objExcel.Cells(row,col+1)

' we need to see if the case has direct deposit for the first warning pop-up
 call navigate_to_MAXIS_screen("MONY", "DISB")
      EMReadScreen DD_Cash, 2, 5, 61   
                 
                'Go to MONY/CHCK
                Call navigate_to_MAXIS_screen ("MONY", "CHCK")
                'You are automatically in edit mode
                                mony_members = 0
                                mony_program_type = "MF"
                                mony_state_code = "MF"
                                mony_reason = "47"      
                                mony_amount = 0           
                                mony_members = total_clients
                                mony_approval = ""
       
                'Writes in the new info
                ' the dates default to March now ...Call Create_MAXIS_friendly_date(date_of_admission_editbox, 0, 4, 43)
                EMWriteScreen mony_program_type, 5, 17
                EMWriteScreen mony_state_code, 5, 21
                EMWriteScreen mony_reason, 5, 32
                'EMWriteScreen mony_amount, 5, 59     ' The amount is given in THE NEXT pop-up for FOOD only
                EMWriteScreen mony_members, 7, 27

                transmit
                
    ' First we will load the amount totals because the age array is looking for _ and will find these spots instead of the correct ones.
row = 1
col = 1
EMSearch "Housing", row, col
'MsgBox ("spot: row: " & row & " col: " & col)
CALL write_value_and_transmit ("xxxxxxxx", row + 3, col + 2)                    ' reset the housing amount so the members don't fill by mistake
CALL write_value_and_transmit ("xxxxxxxx", row + 3, col - 10)                   ' reset cash so the member part does not fill in by mistake
CALL write_value_and_transmit ("xxxxxxxx", row + 3, col + 16)                   ' reset for food issuance so the members don't fill by mistake                            
                
    ' Second we want to read the array and set the first five members as children to bypass the STAT/TIME processing

For EACH ref_num In test_array
   row = 1
   col = 1
   EMSearch "__", row, col
   IF ref_num <> "" and row > 1 THEN
      CALL write_value_and_transmit(ref_num, row, col)
   END IF
NEXT

For EACH age in age_array
   row = 1
   col = 1
   EMSearch "_", row, col
  
  IF age <> "" and row > 1 and row < 15 THEN    ' limit the array to the first page for subsequent pop-ups 
    IF age < 22 then
      CALL write_value_and_transmit("C", row, col)
      CALL write_value_and_transmit("F", row, col + 9)
      CALL write_value_and_transmit("Y", row, col + 19)
    END IF
    IF age = 22 or age > 22 then 
        CALL write_value_and_transmit("C", row, col)
        CALL write_value_and_transmit("F", row, col + 9)
        CALL write_value_and_transmit("Y", row, col + 19)
    END IF 
  END IF
NEXT
   
row = 1
col = 1
EMSearch "Housing", row, col
'MsgBox ("spot: row: " & row & " col: " & col)
CALL write_value_and_transmit ("        ", row + 3, col + 2)                    ' reset the housing amount set on previous screen incorrectly
CALL write_value_and_transmit ("        ", row + 3, col - 10)                   ' reset cash so the member part does not fill in by mistake
CALL write_value_and_transmit ("        ", row + 3, col + 16)                   ' reset the food amount field               
CALL write_value_and_transmit (mfip_amount, row + 3, col + 16)                  ' give the amount from the spreadsheet for food issuance  
  
' we need to test for each pop-up that could appear based on the program written in the pop-up

IF DD_Cash = "DD" then
 Transmit


DO 
   row = 1
   col = 1 
   EMSearch "(Y/N)", row, col
LOOP UNTIL row <> 0

ELSE
DO 
   row = 1
   col = 1 
   EMSearch "(Y/N)", row, col
LOOP UNTIL row <> 0

END IF

'MsgBox (" Where? " & DD_CASE & " ROW: " & row & " COL " & COL)

IF row = 15 then
   IF col = 22 then                                                                          
    
      DO
      	row = 1
            col = 1  
            EMSearch "FMSCIAM9", row, col                                                   ' Please review approval is this correct?
      LOOP UNTIL row <> 0
             
      if row <> 0 then                
      	Do
            	row = 1
                  col = 1
                  EMSearch "Is this correct? (Y/N)  _", row, col
            LOOP UNTIL row <> 0       
		CALL write_value_and_transmit("Y", 15, 29)  
      END IF 
  	DO
  		row = 1
      	col = 1  
      	EMSearch "FMSCIAM2", row, col                                                    ' Final Approval
  	LOOP UNTIL row <> 0
                
   	if row <> 0 then    
     	 	CALL navigate_to_MAXIS_screen("CASE", "NOTE")
     	 	'MsgBox "Checking for privileged"
      	EMReadScreen privileged_case, 40, 24, 2                   'reading any mainframe message on bottom of the map
     	   IF InStr(privileged_case, "PRIVILEGED") = 0 THEN
     		 PF9
             '-----Added PF9 because the script was only case noting the header, footer and worker_signature on the first case.
             '-----Privileged cases will be printed from memo below
             FOR EACH message_part IN message_array
             	CALL write_variable_in_CASE_NOTE(message_part)
                  STATS_counter = STATS_counter + 1    'adds one instance to the stats counter
             NEXT                                                                     
         END IF
                                
         forms_to_arep = ""  'clearing variables otherwise script will try to put a X as variable will remain Y between loops
         forms_to_swkr = ""
         IF MAXIS_case_number <> "" THEN
      	CALL navigate_to_MAXIS_screen("STAT", "MEMB")
            'Checking for privileged
            EMReadScreen privileged_case, 40, 24, 2
            IF InStr(privileged_case, "PRIVILEGED") <> 0 THEN
            	privileged_array = privileged_array & MAXIS_case_number & "~~~"
            ELSE
                  'Navigating to SPEC/MEMO and starting a new memo
                  start_a_new_spec_memo                                          
                  CALL write_variable_in_SPEC_MEMO(memo_text)
                  STATS_counter = STATS_counter + 1    'adds one instance to the stats counter
                  PF4
            END IF
         END IF
      END IF
      transmit
      PF3 
   END IF                  
              
'                                This first spot determines if you have a CLAIM first on ROW 15
IF row = 15 then
   IF col = 49 then
      DO          
     	 row = 1
       col = 1
       EMSearch "FMSCRAM6", row, col                                                  'CLAIMS found
      LOOP UNTIL row <> 0
      if row <> 0 then
         Do
         	row = 1
            col = 1
            EMSearch "Do you want to continue? (Y/N)", row, col                       'CLAIMS QUESTION  Do you want to continue?
         LOOP UNTIL row <> 0 
         EMSetCursor 15,57
         CALL write_value_and_transmit("Y", 15, 57)
      END IF
      DO     
      	row = 1
           	col = 1
            EMSearch "FMSCIAM8", row, col                                             'REI QUESTION
      LOOP UNTIL row <> 0
  
      if row <> 0 then
      	Do
            	row = 1
                  col = 1
                  EMSearch "REI the FOOD portion of the benefit?    (Y/N) _", row, col      
            LOOP UNTIL row <> 0 
            CALL write_value_and_transmit("N", 16, 52)
      END IF
      DO
      	row = 1
            col = 1  
            EMSearch "FMSCIAM9", row, col                                        ' Please review approval is this correct?
      LOOP UNTIL row <> 0
      if row <> 0 then                
      	Do
            	row = 1
                  col = 1
                  EMSearch "Is this correct? (Y/N)  _", row, col
            LOOP UNTIL row <> 0 
            EMSetCursor 15,29
            CALL write_value_and_transmit("Y", 15, 29)  
      END If
      CALL navigate_to_MAXIS_screen("CASE", "NOTE")
      'MsgBox "Checking for privileged"
      EMReadScreen privileged_case, 40, 24, 2                     'reading any mainframe message on bottom of the map
      IF InStr(privileged_case, "PRIVILEGED") = 0 THEN
     		 PF9
             '-----Added PF9 because the script was only case noting the header, footer and worker_signature on the first case.
             '-----Privileged cases will be printed from memo below
             FOR EACH message_part IN message_array
             	CALL write_variable_in_CASE_NOTE(message_part)
                  STATS_counter = STATS_counter + 1    'adds one instance to the stats counter
             NEXT                                                                     
      END IF
                                
      forms_to_arep = ""                                 'clearing variables otherwise script will try to put a X as variable will remain Y between loops
      forms_to_swkr = ""
      IF MAXIS_case_number <> "" THEN
      	CALL navigate_to_MAXIS_screen("STAT", "MEMB")
            'Checking for privileged
            EMReadScreen privileged_case, 40, 24, 2
            IF InStr(privileged_case, "PRIVILEGED") <> 0 THEN
            	privileged_array = privileged_array & MAXIS_case_number & "~~~"
            ELSE
                  'Navigating to SPEC/MEMO and starting a new memo
                  start_a_new_spec_memo                                          
                  CALL write_variable_in_SPEC_MEMO(memo_text)
                  STATS_counter = STATS_counter + 1    'adds one instance to the stats counter
                  PF4
            END IF
       END IF
   END IF
   transmit
   PF3            
END IF 

END IF
 
IF row = 16 then
   IF col = 46 then                                                                          
      DO    
      	row = 1
            col = 1
            EMSearch "FMSCIAM8", row, col                                                   ' REI request
      LOOP UNTIL row <> 0
  'MsgBox ("FMSCIAM8 row: " & row & "FS col: " & col) 
    
      if row <> 0 then
      	Do
			row = 1
                  col = 1
                  EMSearch "REI the FOOD portion of the benefit?    (Y/N) _", row, col       'REI QUESTION for Food?
            LOOP UNTIL row <> 0 
            CALL write_value_and_transmit("N", 16, 52)
      END IF
      DO
      	row = 1
            col = 1  
            EMSearch "FMSCIAM9", row, col                                                   ' Please review approval is this correct?
      LOOP UNTIL row <> 0
             
      if row <> 0 then                
      	Do
            	row = 1
                  col = 1
                  EMSearch "Is this correct? (Y/N)  _", row, col
            LOOP UNTIL row <> 0       
		CALL write_value_and_transmit("Y", 15, 29)  
      END IF 
  	DO
  		row = 1
      	col = 1  
      	EMSearch "FMSCIAM2", row, col                                                    ' Final Approval
  	LOOP UNTIL row <> 0
                
   	if row <> 0 then    
     	 	CALL navigate_to_MAXIS_screen("CASE", "NOTE")
     	 	'MsgBox "Checking for privileged"
      	EMReadScreen privileged_case, 40, 24, 2                   'reading any mainframe message on bottom of the map
     	   IF InStr(privileged_case, "PRIVILEGED") = 0 THEN
     		 PF9
             '-----Added PF9 because the script was only case noting the header, footer and worker_signature on the first case.
             '-----Privileged cases will be printed from memo below
             FOR EACH message_part IN message_array
             	CALL write_variable_in_CASE_NOTE(message_part)
                  STATS_counter = STATS_counter + 1    'adds one instance to the stats counter
             NEXT                                                                     
         END IF
                                
         forms_to_arep = ""  'clearing variables otherwise script will try to put a X as variable will remain Y between loops
         forms_to_swkr = ""
         IF MAXIS_case_number <> "" THEN
      	CALL navigate_to_MAXIS_screen("STAT", "MEMB")
            'Checking for privileged
            EMReadScreen privileged_case, 40, 24, 2
            IF InStr(privileged_case, "PRIVILEGED") <> 0 THEN
            	privileged_array = privileged_array & MAXIS_case_number & "~~~"
            ELSE
                  'Navigating to SPEC/MEMO and starting a new memo
                  start_a_new_spec_memo                                          
                  CALL write_variable_in_SPEC_MEMO(memo_text)
                  STATS_counter = STATS_counter + 1    'adds one instance to the stats counter
                  PF4
            END IF
         END IF
      END IF
      transmit
      PF3 
   END IF                  
END IF               
                
   
                                                                
END FUNCTION



'The script===========================
EMConnect ""

CALL check_for_MAXIS(true)
copy_case_note = FALSE
run_mode = "Excel File"

'>>>>> loading the main dialog <<<<<
'The main dialog was modified to request the benefit date range with the Excel Spreadsheet as the default

DIALOG main_menu
                IF ButtonPressed = 0 THEN stopscript
                '>>>>> the script has different ways of building case_number_array
                IF run_mode = "Manual Entry" THEN
                   CALL build_manual_entry_dlg(case_number_array, case_note_header, case_note_body, worker_signature)
                ELSEIF run_mode = "Excel File" THEN
                                'Opening the Excel file

                                DO
                                                call file_selection_system_dialog(excel_file_path, ".xlsx")

                                                Set objExcel = CreateObject("Excel.Application")
                                                Set objWorkbook = objExcel.Workbooks.Open(excel_file_path)
                                                objExcel.Visible = True
                                                objExcel.DisplayAlerts = True
                                                                                                
                                                confirm_file = MsgBox("Is this the correct file? Press YES to continue. Press NO to try again. Press CANCEL to stop the script.", vbYesNoCancel)
                                                IF confirm_file = vbCancel THEN
                                                                objWorkbook.Close
                                                                objExcel.Quit
                                                                stopscript
                                                ELSEIF confirm_file = vbNo THEN
                                                                objWorkbook.Close
                                                                objExcel.Quit
                                                END IF
                                LOOP UNTIL confirm_file = vbYes

                                'Gathering the information from the user about the fields in Excel to look for.
                                DO
                                                err_msg = ""
                                                DIALOG CASE_NOTE_from_excel_dlg
                                                                IF ButtonPressed = 0 THEN stopscript
                                                                IF isnumeric(excel_col) = FALSE AND len(excel_col) > 2 THEN
                                                                                err_msg = err_msg & vbCr & "* Please do not use such a large column. The script cannot handle it."
                                                                ELSE
                                                                                IF (isnumeric(right(excel_col, 1)) = TRUE AND isnumeric(left(excel_col, 1)) = FALSE) OR (isnumeric(right(excel_col, 1)) = FALSE AND isnumeric(left(excel_col, 1)) = TRUE) THEN
                                                                                                err_msg = err_msg & vbCr & "* Please use a valid Column indicator. " & excel_col & " contains BOTH a letter and a number."
                                                                                ELSE
                                                                                                call convert_excel_letter_to_excel_number(excel_col)
                                                                                                IF isnumeric(excel_row) = false or isnumeric(end_row) = false THEN err_msg = err_msg & vbCr & "* Please enter the Excel rows as numeric characters."
                                                                                                IF end_row = "" THEN err_msg = err_msg & vbCr & "* Please enter an end to the search. The script needs to know when to stop searching."
                                                                                END IF
                                                                END IF
                                                                IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
                                LOOP UNTIL err_msg = ""
                                

                                CALL check_for_MAXIS(false)
                                'Generating a CASE NOTE for each case.
                                FOR i = excel_row TO end_row
                                                IF objExcel.Cells(i, excel_col).Value <> "" THEN
                                                                case_number_array = case_number_array & objExcel.Cells(i, excel_col).Value & "~~~"
                                                                mfip_add_array = mfip_add_array & objExcel.Cells(i, excel_col+1).Value  & "$$"     'note: the "$$" is a separator of the values
                                                END IF
                                NEXT
                END IF

CALL check_for_MAXIS(false)

'The business of sending Case notes
case_number_array = trim(case_number_array)
case_number_array = split(case_number_array, "~~~")
mfip_add_array = trim(mfip_add_array)
mfip_add_array = split(mfip_add_array, "$$")


'Formatting case note
If copy_case_note = FALSE Then
                message_array = case_note_header & "~%~" & case_note_body & "~%~" & "---" & "~%~" & worker_signature & "~%~" & "---" & "~%~" & "**Processed in bulk script**"
                message_array = split(message_array, "~%~")
End If

privileged_array = ""

CALL check_for_MAXIS(false)

FOR EACH MAXIS_case_number IN case_number_array
                IF MAXIS_case_number <> "" THEN
                               
                                call create_mony_chck_mf ()


   END IF
NEXT

IF privileged_array <> "" THEN
                privileged_array = replace(privileged_array, "~~~", vbCr)
                MsgBox "The script could not generate a NOTE/memo for the following cases..." & vbCr & privileged_array
END IF

STATS_counter = STATS_counter - 1  'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success!!")