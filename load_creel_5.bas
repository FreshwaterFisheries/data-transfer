Attribute VB_Name = "load_creel_5"
Option Compare Database
Private Sub load_creel()
'loads Creel data into creel_survey
'includes a corresponding entry in assessment_id
'links to survey_type_id
'populates creel_fish_counts

On Error GoTo load_creel_Err

Dim db As DAO.Database
Dim rsData As DAO.Recordset             'data to load
Dim rsCS As New ADODB.Recordset         'creel_survey table
Dim rsFC As New ADODB.Recordset         'creel_fish_count table
Dim rsCArch As New ADODB.Recordset      'creel_archived table
Dim rsSpCaught As New ADODB.Recordset   'species lookup table
Dim rsAsmnt As New ADODB.Recordset      'assessment table
Dim rsProj_Asmnt As New ADODB.Recordset 'project-assessment table
Dim rsTemp As New ADODB.Recordset       'temporary recordset
Dim conn As New ADODB.Connection        'connection to small lakes

Dim cmd As New ADODB.Command
Dim match As Boolean

Dim a_id As Variant
Dim cs_id As Integer
Dim s As String
Dim sp_code As String
Dim sp_id As Integer
Dim target_sp1_id As Integer
Dim target_sp2_id As Integer
Dim ag_id As Integer
Dim lt_id As Integer

outfilepath = "\\FFSCOFP04\Users$\bonnie.robert\My Documents\Projects\DB Restructuring\Logfiles\creel_load_log.txt"
crld_log = FreeFile()
Close #crld_log
Open outfilepath For Output As #crld_log
Print #crld_log, "Starting QC error report: " & Now

'open connection to small lakes test database
conn.Open "SMALL_LAKES-TEST", "GOFISHBC/Bonnie.Robert"

'open Creel table in local Access instance
Set db = CurrentDb
Set rsData = db.OpenRecordset("SELECT * FROM ffsbc_Creel;")

'open fish_count table for adding new records
rsFC.CursorType = adOpenKeyset
rsFC.LockType = adLockOptimistic
rsFC.Open "SELECT * FROM ffsbc.creel_fish_count;", conn

'open creel_survey table to add new records
rsCS.CursorType = adOpenKeyset
rsCS.LockType = adLockOptimistic
rsCS.Open "SELECT * FROM ffsbc.creel_survey;", conn

'open creel_archived table to add new records
rsCArch.CursorType = adOpenKeyset
rsCArch.LockType = adLockOptimistic
rsCArch.Open "SELECT * FROM ffsbc.creel_archived;", conn

'open project_assessment table table to add new records
rsProj_Asmnt.CursorType = adOpenKeyset
rsProj_Asmnt.LockType = adLockOptimistic
rsProj_Asmnt.Open "SELECT * FROM ffsbc.project_assessment;", conn

'open assessment table to new add records
rsAsmnt.CursorType = adOpenKeyset
rsAsmnt.LockType = adLockOptimistic
rsAsmnt.Open "SELECT * FROM ffsbc.assessment;", conn

'determine which species were caught and recorded in the creel data sheet
For Each fld In rsData.Fields
    If InStr(fld.Name, "_Caught") Then
        s = s + " " + Left(fld.Name, InStr(fld.Name, "_Caught") - 1)
    End If
Next fld
s = Replace(s, " ", "', '", 2)

'open a recordset with species codes that were caught and recorded in the Creel data sheet
rsSpCaught.Open "SELECT * FROM ffsbc.species WHERE species_code IN ('" & s & "');", conn

'for each record in creel dataset (read only)
Do Until rsData.EOF
    Debug.Print ("Processing Creel_ID: " & rsData.Fields("Creel_ID"))
    
'open assessment table to see if there's an associated assessment
    a_key = rsData.Fields("Assessment_ID")
    
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.assessment WHERE assessment_key = '" & a_key & "';", conn
    
    If (rsTemp.RecordCount = 0) Then
        a_id = 0
    Else:
        a_id = rsTemp.Fields("assessment_id")
    End If
    rsTemp.Close
  
 'open project table to get related project id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT project_id FROM ffsbc.project WHERE project_name = '" & Trim(rsData.Fields("FFSBC_ID")) & "';", conn
                
    If (rsTemp.RecordCount = 0) Then
        p_id = 88
    Else:
        p_id = rsTemp.Fields("project_id")
    End If
    rsTemp.Close
  
    'open species to get related species id  for target species 1
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.species WHERE species_code = '" & Trim(rsData.Fields("target_species")) & "';", conn
    
    If (rsTemp.RecordCount = 0) Then
        target_sp1_id = 62
    Else:
        target_sp1_id = rsTemp.Fields("species_id")
    End If
    rsTemp.Close

   'open lookup_age_group to get related age_group id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.lookup_age_group WHERE age_group = '" & Trim(rsData.Fields("Age_Group")) & "';", conn
    
    If (rsTemp.RecordCount = 0) Then
        ag_id = Null
    Else:
        ag_id = rsTemp.Fields("lookup_age_group_id")
    End If
    rsTemp.Close
    
    'open lookup_licence_type to get related lookup_licence_type
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.lookup_licence_type WHERE licence_type = '" & Trim(rsData.Fields("Type_Of_License")) & "';", conn
    
    If (rsTemp.RecordCount = 0) Then
        lt_id = Null
    Else:
        lt_id = rsTemp.Fields("lookup_licence_type_id")
    End If
    rsTemp.Close

    'open waterbody table to get related waterbody id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.waterbody WHERE MOF_waterbody_id = '" & Trim(rsData.Fields("WBID")) & "';", conn
                      
    If (rsTemp.RecordCount = 0) Then
        wb_id = 0
    Else:
        wb_id = rsTemp.Fields("waterbody_id")
    End If
    rsTemp.Close
    
    'open region table to get related region id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.waterbody_region WHERE waterbody_id = '" & wb_id & "';", conn
                      
    If (rsTemp.RecordCount = 0) Then
        r_id = 13
    Else:
        r_id = rsTemp.Fields("region_id")
    End If
    rsTemp.Close

    'add assessment and assessment/project record if none in assessment table
    If (a_id = 0) Then
        With rsAsmnt
            .AddNew
                !waterbody_id = wb_id
                !Assessment_Key = a_key
                !region_id = r_id
                !Source = rsData.Fields("Source")
                !start_date = rsData.Fields("Date")
                !end_date = rsData.Fields("Date")
                !lookup_method_id = 4
                !date_added = Date
                !comments = "Creel transferred from old SL database table 4 - Creel"
                !date_updated = Date
            .Update
            a_id = rsAsmnt.Fields("assessment_id")
        End With
        
        'add assessment_id and project_id to assessment_project table
        With rsProj_Asmnt
            .AddNew
                !project_id = p_id
                !assessment_id = a_id
            .Update
        End With
        
     End If
    
    'create new record in creel_survey and enter data for current record
    With rsCS
            .AddNew
            !assessment_id = a_id
            !lookup_method_id = 4
            !date_of_interview = rsData.Fields("Date")
            !time_of_interview = rsData.Fields("Time_of_Interview")
            !launch_or_fishing_location = rsData.Fields("Lake_Location")
            !boat_shore = rsData.Fields("Boat_Type")
            !party_no = rsData.Fields("Party_No")
            !Person_No = rsData.Fields("Person_No")
            !no_anglers = rsData.Fields("No_in_Party")
            !no_rods = rsData.Fields("No_Rods")
         '' !target_species = rsData.Fields("Target_Species")
            !lookup_target_species_id = target_sp1_id
         '' !lookup_target_species_id = target_sp2_id
            !hours_fished = rsData.Fields("hours_fished")
            !no_expected_to_catch = rsData.Fields("No_Expected_to_Catch")
            !no_expected_to_keep = rsData.Fields("No_Expected_to_Keep")
            !length_weight_measured = rsData.Fields("Length_Weight_Measured")
            !fishing_experience_today = rsData.Fields("Fishing_Experience_Today")
            !days_freshwater_BC_current_year = rsData.Fields("Days_Freshwater_BC_Current_Year")
            !years_freshwater_fishing_BC = rsData.Fields("Years_Freshwater_Fishing_BC")
            !days_freshwater_BC_prev_two_years = rsData.Fields("Days_Freshwater_BC_Prev_Two_Years")
            !days_ice_BC_prev_two_years = rsData.Fields("Days_Ice_BC_Prev_Two_Years")
            !no_of_licences_purchased_in_last_5_years = rsData.Fields("No_of_licenses_purchased_in_last_5_years")
            !previous_survey_this_year = rsData.Fields("Previous_Survey_This_Year")
            !type_of_licence = rsData.Fields("Type_of_License")
            !age_group = rsData.Fields("Age_Group")
            !age = rsData.Fields("Age")
            !gender = rsData.Fields("Gender")
            !hometown_postal_code_first_3_digits = rsData.Fields("Hometown_Postal_Code_First_3_Digits")
            !mother_country = rsData.Fields("Mother_Country")
            !Phone_Number_Last_4_Digits = rsData.Fields("Phone_Number_Last_4_Digits")
            !time_travelled_to_lake = rsData.Fields("Time_Travelled_to_Lake")
            !aware_of_stocking = rsData.Fields("Aware_of_stocking")
            !terminal_gear_type = rsData.Fields("Terminal_Gear_Type")
            !interviewer = rsData.Fields("Interviewer")
            !days_freshwater_ice_salt_last_2_years = rsData.Fields("Days_Freshwater_Ice_Salt_last_2_years")
            !percent_days_fished_KO = rsData.Fields("Percent_Days_Fished_KO")
            !Boat_Type = rsData.Fields("Boat_Type")
            !percent_increase_in_fishing_days_past_2_years = rsData.Fields("Percent_Increase_in_fishing_days_past_2_years")
            !how_many_days_were_overnight_fishing_trips_in_2_years = rsData.Fields("How_many_days_were_overnight_fishing_trips_in_2_years")
            !where_staying_tonight = rsData.Fields("Where_staying_tonight")
            !how_many_days_in_current_trip = rsData.Fields("How_many_days_in_current_trip")
            !how_many_days_of_trip_will_you_fish_this_lake_or_other_lakes = rsData.Fields("How_many_days_of_trip_will_you_fish_this_lake_or_other_lakes")
            !prefer_to_catch_lots_small_1_or_few_large_7 = rsData.Fields("Prefer_to_catch_lots_smal_1_or_few_large_7")
            !cost_gas_for_this_trip = rsData.Fields("Cost_Gas_for_this_trip")
            !cost_lodging_for_this_trip = rsData.Fields("Cost_Lodging_for_this_trip")
            !cost_food_for_this_trip = rsData.Fields("Cost_Food_for_this_trip")
            !spend_annually_on_tackle_rods_or_reels = rsData.Fields("Spend_annually_on_tackle_rods_or_reels")
            !spend_annually_on_boats_motors_or_equipment = rsData.Fields("Spend_annually_on_boats_motors_or_equipment")
            !motivation_method = rsData.Fields("Motivation_Method")
            !motivation_to_catch_fish_for_eating = rsData.Fields("Motivation_To_Catch_Fish_for_Eating")
            !motivation_to_catch_large_fish = rsData.Fields("Motivation_To_Catch_Large_Fish")
            !motivation_to_catch_many_fish = rsData.Fields("Motivation_To_Catch_Many_Fish")
            !motivation_challenge = rsData.Fields("Motivation_Challenge")
            !motivation_to_get_away = rsData.Fields("Motivation_To_Get_Away")
            !motivation_relaxation = rsData.Fields("Motivation_Relaxation")
            !motivation_family_closer_together = rsData.Fields("Motivation_Family_Closer_Together")
            !motivation_companionship = rsData.Fields("Motivation_Companionship")
            !motivation_improve_skills = rsData.Fields("Motivation_Improve_Skills")
            !motivation_close_to_nature = rsData.Fields("Motivation_Close_to_Nature")
            !top_waterbodies_waterbody_1 = rsData.Fields("Top_Waterbodies_Waterbody_1")
            !top_waterbodies_rank_1 = rsData.Fields("Top_Waterbodies_Rank_1")
            !top_waterbodies_waterbody_2 = rsData.Fields("Top_Waterbodies_Waterbody_2")
            !top_waterbodies_rank_2 = rsData.Fields("Top_Waterbodies_Rank_2")
            !top_waterbodies_waterbody_3 = rsData.Fields("Top_Waterbodies_Waterbody_3")
            !top_waterbodies_rank_3 = rsData.Fields("Top_Waterbodies_Rank_3")
            !date_added = Date
            !comments = rsData.Fields("Comments")
            .Update
            
            'get primary key to enter in creel_fish_counts
            cs_id = rsCS.Fields("creel_survey_id")
            
    End With
    
    'create new record in creel_archived and enter archived data for current record
    With rsCArch
    
        .AddNew
        !creel_survey_id = cs_id
        !Recreation_1_vs_Food_5 = rsData.Fields("Recreation_1_vs_Food_5")
        !Days_Freshwater_Last_Year = rsData.Fields("Days_Freshwater_Last_Year")
        !Days_Freshwater_Last_Two_Years = rsData.Fields("Days_Freshwater_Last_Two_Years")
        !Days_Saltwater_Each_Year = rsData.Fields("Days_Saltwater_Each_Year")
        !Days_Saltwater_Last_Two_Years = rsData.Fields("Days_Saltwater_Last_Two_Years")
        !Days_Ice_Fishing_This_Year = rsData.Fields("Days_Ice_Fishing_This_Year")
        !Days_Ice_Fishing_Each_Year = rsData.Fields("Days_Ice_Fishing_Each_Year")
        !Days_Fished_This_Lake_This_Year = rsData.Fields("Days_Fished_This_Lake_This_Year")
        !Angler_Type_Days_Fly_Casting = rsData.Fields("Angler_Type_Days_Fly_Casting")
        !Angler_Type_Days_Spin_Casting = rsData.Fields("Angler_Type_Days_Spin_Casting")
        !Angler_Type_Days_Trolling = rsData.Fields("Angler_Type_Days_Trolling")
        !terminal_gear_days_artificial_fly = rsData.Fields("Terminal_Gear_Days_Artifical_Fly")
        !Terminal_Gear_Days_Lure = rsData.Fields("Terminal_Gear_Days_Lure")
        !termina_gear_days_bait = rsData.Fields("Termina_Gear_Day_s_Bait")
        !date_added = Date
        .Update
    
    End With

    'set current record in species caught recordset to beginning record
    rsSpCaught.MoveFirst

    'for each species found in the Creel table
    Do Until rsSpCaught.EOF

        sp_code = Trim(rsSpCaught.Fields("species_code"))
        sp_id = rsSpCaught.Fields("species_id")
        
        'if there is data (i.e. not 0 or NULL), create a record in creel_fish_count
        If rsData.Fields(sp_code & "_Caught") Or rsData.Fields(sp_code & "_Kept") Then
        
            With rsFC
                .AddNew
                !number_caught = rsData.Fields(sp_code & "_Caught")
                !number_kept = rsData.Fields(sp_code & "_Kept")
                !species_id = sp_id
                !creel_survey_id = cs_id
                .Update
            End With
        End If
        
        'next species found
        rsSpCaught.MoveNext
    Loop


'*****************************finished processing current data record
'set current record in Creel data to next record
rsData.MoveNext
Loop

'close handles to commit changes
rsData.Close
rsCS.Close
rsSpCaught.Close
rsFC.Close
rsCArch.Close
rsAsmnt.Close
rsAsmnt_Proj.Close
conn.Close

Exit_load_creel:
    DoCmd.SetWarnings True
    Exit Sub

load_creel_Err:
    MsgBox Err.Description
    Resume Exit_load_creel
End Sub




