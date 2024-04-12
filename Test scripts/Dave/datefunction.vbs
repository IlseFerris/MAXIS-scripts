'Required for statistical purposes==========================================================================================
name_of_script = "Test.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 500                     'manual run time in seconds
STATS_denomination = "C"                   'C is for each CASE
'END OF stats block================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
    IF on_the_desert_island = TRUE Then
        FuncLib_URL = "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Script Files\desert-island\MASTER FUNCTIONS LIBRARY.vbs"
        Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
        Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
        text_from_the_other_script = fso_command.ReadAll
        fso_command.Close
        Execute text_from_the_other_script
    ELSEIF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message
			critical_error_msgbox = MsgBox ("Something has gone wrong. The Functions Library code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine &_
                                            "FuncLib URL: " & FuncLib_URL & vbNewLine & vbNewLine &_
                                            "The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", _
                                            vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
            StopScript
		END IF
	ELSE
		FuncLib_URL = "C:\MAXIS-scripts\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF


'This class is necessary for the HH_member_enhanced_dialog. 
	Class member_data
		public member_number
		public name
		public ssn
		public birthdate
        public PMI_number
        public first_checkbox
        public second_checkbox
        public elig_type
	End Class


function HH_member_enhanced_dialog(HH_member_array, instruction_text, display_birthdate, display_ssn, first_checkbox, first_checkbox_default, second_checkbox, second_checkbox_default)
'--- This function creates an array of all household members in a MAXIS case, and displays a dialog of HH members that allows the user to select up to two checkboxes per member.
'~~~~~ enhanced_HH_member_array: array that stores all members of the household, with attributes for each member stored in an object. 
'~~~~~ instruction_text: String variable that will appear at the top of dialog as text to give instructions or other info to the user. Limit to 400 characters????
'~~~~~ display_birthdate: true/false. True will display the birthdate after the member name for each HH member
'~~~~~ display_ssn: true/False. True will display the last 4 digits of the SSN after the member name for each HH member
'~~~~~ first_checkbox: string value that contains the text to display for the first checkbox. If no checkbox is wanted, set to ""
'~~~~~ first_checkbox_default: checked/unchecked or 0/1. Determines default state of first checkbox.
'~~~~~ second_checkbox: string value that contains the text to display for the second checkbox. If no checkbox is wanted, set to ""
'~~~~~ second_checkbox_default: checked/unchecked or 0/1. Determines default state of first checkbox.
'If both checkboxes are set to "", the dialog will not display. Use this option when populating an array of the whole household.
'The 6 attributes (member_number, name, ssn, birthdate, first_checkbox, second_checkbox) will be stored in the enhanced_hh_member_array and can be called with this syntax: enhanced_hh_member_array(member).birthdate 
'===== Keywords: MAXIS, member, array, dialog
dim enhanced_HH_member_array()
	call check_for_MAXIS(false)
	membs = 1
    'redim enhanced_HH_member_array(1)
	CALL Navigate_to_MAXIS_screen("STAT", "MEMB")   'navigating to stat memb to gather the ref number and name.
	EMWriteScreen "01", 20, 76						''make sure to start at Memb 01
    transmit
	EMREadScreen total_clients, 2, 2, 78
	total_clients = cint(replace(total_clients, " ", ""))
	redim enhanced_HH_member_array(total_clients, 6)
	DO								'reads the reference number, last name, first name, and then puts it into a single string then into the array
		 'if only one MEMB screen, we don't need to display the dialog 
		
        EMReadScreen access_denied_check, 13, 24, 2
        'MsgBox access_denied_check
        If access_denied_check = "ACCESS DENIED" Then
            PF10
			EMWaitReady 0, 0
            last_name = "UNABLE TO FIND"
            first_name = " - Access Denied"
            mid_initial = ""
			ssn_last_4 = ""
			birthdate = ""
        Else
            EMReadscreen ref_nbr, 3, 4, 33
    		EMReadscreen last_name, 25, 6, 30
    		EMReadscreen first_name, 12, 6, 63
    		EMReadscreen mid_initial, 1, 6, 79
			EMReadScreen ssn, 11, 7, 42 
			EMReadScreen birthdate, 10, 8, 42
            EMReadScreen PMI_number, 8, 4, 46
    		last_name = trim(replace(last_name, "_", "")) & " "
    		first_name = trim(replace(first_name, "_", "")) & " "
    		mid_initial = replace(mid_initial, "_", "")
			birthdate = replace(birthdate, " ", "/")
		End If
		client_string = last_name & first_name & mid_initial
		'Create an object for the member and add that members info, plus the checkbox defaults
        		
		enhanced_HH_member_array(membs, 0) = ref_nbr
		enhanced_HH_member_array(membs, 1) = client_string
		enhanced_HH_member_array(membs, 2) = replace(ssn, " ", "") 
		enhanced_HH_member_array(membs, 3) = birthdate
		enhanced_HH_member_array(membs, 4) = first_checkbox_default
		enhanced_HH_member_array(membs, 5) = second_checkbox_default
        enhanced_HH_member_array(membs, 6) = PMI_number
      

  		membs = membs + 1 'index the value up 1 for next member
		transmit
	    Emreadscreen edit_check, 7, 24, 2

	LOOP until edit_check = "ENTER A"			'the script will continue to transmit through memb until it reaches the last page and finds the ENTER A edit on the bottom row.

	instruction_text_lines = (len(instruction_text) \ 80) + 1
	if total_clients > 8 Then instruction_text_lines = (len(instruction_text) \ 160) + 1
	If first_checkbox <> "" OR second_checkbox <> "" Then 
		If total_clients > 1 OR second_checkbox <> "" Then 'We only need the dialog if more than 1 client or multiple checkboxes to select
			'Generating the dialog
			split_number = 9
			If total_clients > 8 Then split_number = (total_clients \ 2) + 1
		    member_height = 15
		    If display_ssn = true Or display_birthdate = true  Then member_height = member_height + 15
		    If first_checkbox <> "" Then member_height = member_height + 15
		    If second_checkbox <> "" Then member_height = member_height + 15

		    If total_clients < split_number Then 'Single column dialog
				dialog_width = 290
				dialog_height = (total_clients * 35) + (instruction_text_lines * 15) + 20
			Else
				dialog_width = 580
				dialog_height = (split_number * 35) + (instruction_text_lines * 15) + 20
			End If 
			dialog1 = ""

		    BEGINDIALOG dialog1, 0, 0, dialog_width, dialog_height, "HH Member Dialog"   
				y_pos = 5
		    	Text 10, y_pos, dialog_width - 20, 10 * instruction_text_lines, instruction_text
				y_pos = y_pos + (10 * instruction_text_lines) + 10

		    	FOR person = 1 to total_clients										
		    		'enhanced_HH_member_array(i).member_number
    	            x_pos = 10
					IF enhanced_HH_member_array(person, 0) <> "" THEN 
		    			if person > split_number THEN x_pos = 300
						display_string = enhanced_HH_member_array(person, 1) 'client name
						If display_birthdate = True Then display_string = display_string & " " & enhanced_HH_member_array(person, 3) 'birthdate
						If display_ssn = True Then display_string = display_string & "  XXX-XX-" & right(enhanced_HH_member_array(person, 2), 4) 'ssn
						Text x_pos, y_pos, 270, 10, enhanced_HH_member_array(person, 0) & " " & display_string   'Ignores and blank scanned in persons/strings to avoid a blank checkbox
		    			If first_checkbox <> "" Then checkbox x_pos + 10, y_pos + 15, 125, 10, first_checkbox, enhanced_HH_member_array(person, 4) 'enhanced_HH_member_array(i).first_checkbox 
    	                If second_checkbox <> "" Then checkbox x_pos + 140, y_pos + 15, 125, 10, second_checkbox, enhanced_HH_member_array(person, 5)   
		    			y_pos = y_pos + 30
						if person = split_number Then y_pos = 15 + (10 * instruction_text_lines) 'resets y value when moving to next column
    	            End If
		    	NEXT
		    	ButtonGroup ButtonPressed
		    	OkButton dialog_width - 115, dialog_height - 20, 50, 15
		    	CancelButton dialog_width - 60, dialog_height - 20, 50, 15 
		    ENDDIALOG
			'runs the dialog that has been dynamically created
	
    		'Put this in a loop and make sure they actually check something
		    Dialog dialog1
		    Cancel_without_confirmation
		End If 
	End If 
	'This section puts each person's info into objects in teh HH_member_array
	redim hh_member_array(total_clients)
	For memb = 1 to total_clients
		set HH_member_array(memb) = new member_data
		with HH_member_array(memb)
			.member_number = enhanced_HH_member_array(memb, 0)
			.name = enhanced_HH_member_array(memb, 1)
			.ssn = enhanced_HH_member_array(memb, 2)
			.birthdate = enhanced_HH_member_array(memb, 3)
			.first_checkbox = enhanced_HH_member_array(memb, 4)
			.second_checkbox = enhanced_HH_member_array(memb, 5)
            .PMI_number = enhanced_HH_member_array(memb, 6)
		end with
	next
end function
 Function display_exemptions() 'A message box showing exemptions from SNAP work rules
	wreg_exemptions = msgbox("Individuals in your household may not have to follow these General Work Rules if [you/they] are:" & vbCr & vbCr &_
				"* Explain to the resident which members of the household are subject to the work rules. *" & vbCr &_
	     		  "* Younger than 16 or older than 59," & vbCr &_
	     		  "* Taking care of a child younger than 6 or someone who needs helps caring for themselves, " & vbCr &_
	     		  "* Already working at least 30 hours a week," & vbCr &_
	     		  "* Already earning $217.50 or more per week," & vbCr &_
	     		  "* Receiving unemployment benefits, or you applied for unemployment benefits," & vbCr &_
	     		  "* Not working because of a physical illness, injury, disability, or surgery recovery," & vbCr &_
	     		  "* Not working due to a mental health illness, disorder, or health condition," & vbCr &_
				  "* Are homeless," & vbCr &_
				  "* A victim of domestic violence," & vbCr &_
				  "* Going to school, college, or a training program at least half time," & vbCr &_
				  "* Meeting the work rules for Minnesota Family Investment Program (MFIP) or DWP (Divisionary Work Program (DWP)," & vbCr &_
				  "* Not working due to a substance use disorder or addiction dependency, or" & vbCr &_
				  "* Participating in a drug or alcohol addiction treatment program." & vbCr & vbCr &_
				  "Press yes if you reviewed exemptions with the resident, press no to return to the previous dialog without review." & vbCr &_
				  "Press 'Cancel' to end the script run.", vbYesNoCancel+ vbQuestion, "Work Rules Reviewed")
		If wreg_exemptions = vbCancel then cancel_confirmation
	If wreg_exemptions = vbYes then work_exemptions_reviewed = true
End Function
Function display_work_rules(work_rules_members) 'displays a dialog showing the general work rules for SNAP, including a list of members that may be subject to rules
	Call HH_member_enhanced_dialog(HH_member_array, "", false, false, "", false, "", false) 'This collects the birthdates and other data for all hh members without displaying the dialog
		'Establish the birthdates for the age cutoffs, which are the 1st of the month after the birthday for given age. want under 16 or over 59
		month_date = cdate(datepart("m", dateadd("m", 1, date)) & "/01/" & datepart("yyyy", dateadd("m", 1, date)))
		sixteen_date = dateadd("yyyy", -16, month_date) 
		fity_nine_date = dateadd("yyyy", -59, month_date) 
	'Go through the member_array we just generated and check ages
		For memb = 1 to ubound(HH_member_array)
			if isdate(HH_member_array(memb).birthdate) = True Then
				if HH_member_array(memb).birthdate < sixteen_date AND HH_member_array(memb).birthdate > fifty_nine_date Then HH_member_array(memb).first_checkbox = True 'using first checkbox to store the users we need
			Else
				HH_member_array(memb).first_checkbox = true 'If we don't have a birthdate for this memb, put them in the list
			End if 
		Next
			dim memb_list()
			BeginDialog Dialog1, 0, 0, 385, 300, "SNAP General Work Rules"
				 Text 15, 25, 350, 20, "First, explain to the resident which members of the household are subject to the work rules, and select those members below."
				 Text 15, 50, 350, 10, "		          -------------------------------------------------------------------------------------					"
				 y_pos = 50
				 memb_count = 1
				 'create a checkbox for each member that is not age exempt.They get put in a 2-d array, column 2 element 1 is the checkbox value, element 2 is the member's object
				 For memb = 1 to ubound(HH_member_array)
				 	
					If HH_member_array(memb).first_checkbox = True Then
						msgbox HH_member_array(memb).name
						if memb_count mod 2 = 1 then  'putting odd numbers in left column
							x_loc = 20
							y_pos = y_pos + 15
						else 'even in right
							x_loc = 250
						End if 
						redim preserve memb_list(memb_count, 2)
						checkbox, x_loc, y_pos, 200, 10, HH_member_array(memb).name, memb_list(memb_count, 1)
						memb_count = memb_count + 1
					End If 
				 Next
				 Text 15, y_pos + 55, 350, 10, ""
	     		 Text 15, y_pos + 70, 350, 10, "To follow the general work rules, these members must:"
	     		 Text 15, y_pos + 85, 350, 10, "* Accept any job offer received, unless there is a good reason they can't. "
	     		 Text 15, y_pos + 100, 350, 20, "* If they have a job, don't quit or choose to work less than 30 hours each week without having a good reason. Good reasons could be getting sick, being discriminated against, or not getting paid."
	     		 Text 15, y_pos + 125, 350, 10, "* Tell us about your job and how much you are working, if asked."
	     		 Text 15, y_pos + 140, 350, 10, "* You may lose your SNAP benefits if you don't follow these work rules without having a good reason."
	     		 Text 15, y_pos + 155, 350, 10, "It is important for you to know that there are consequences if you/they don't follow these General Work Rules: "
	     		 Text 15, y_pos + 170, 350, 20, "The first time [you/they] don't follow these rules, and you don't have a good reason, you can't get SNAP benefits for 1 month."
				 Text 15, y_pos + 195, 350, 10, "The second time [you/they] don't follow these rules, you can't get SNAP benefits for 3 months."
				 Text 15, y_pos + 210, 350, 10, "The third time, and any time after that, [you/they] can't get SNAP benefits for 6 months."
				 Text 15, y_pos + 225, 350, 10, "		          -------------------------------------------------------------------------------------					"
				ButtonGroup ButtonPressed
				 PushButton 20, 260, 145, 15, "Press here to review a list of exemptions.", exemptions_button
  				 PushButton 210, 240, 145, 15, "Press here to continue without reviewing.", continue_button
  				 PushButton 20, 240, 145, 15, "Press here if you reviewed with resident.", work_rules_reviewed_button
				 PushButton 210, 260, 145, 15, "Press here to return to the previous dialog.", return_to_info_btn
			EndDialog
			'Display the dialog
			Do 
				Dialog Dialog1
			Loop until ButtonPressed = continue_button
			work_rules_members = ""
			For i = 0 to ubound(memb_list, 1)
				If memb_list(i, 1) = checked Then 
					work_rules_members = work_rules_members & "Memb: " & memb_list(i, 2).member_number & " " & memb_list(i,2).name & ", "
				End if 
			Next 
			work_rules_members = left(work_rules_members, len(work_rules_members)-2) 'this pulls the right two chars off the string to ditch the ", "
End Function
check_for_MAXIS(false)
call MAXIS_case_number_finder(MAXIS_case_number)
call display_work_rules(work_rules_members) 
stopscript