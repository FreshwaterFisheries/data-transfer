Attribute VB_Name = "load_biological_data_8"
Option Compare Database
Private Sub load_biological_data()
'loads biological_data data into biological_data_survey
'includes a corresponding entry in assessment_id
'links to survey_type_id
'populates biological_data_fish_counts

On Error GoTo load_biological_data_Err

'declare recordset and connection variables for data transfer
Dim db As DAO.Database
Dim rsData As DAO.Recordset             'data to load
Dim rsIF As New ADODB.Recordset         'individual fish table
Dim rsCS As New ADODB.Recordset         'creel survey table
Dim rsAsmnt As New ADODB.Recordset      'assessment table
Dim rsProj_Asmnt As New ADODB.Recordset 'project-assessment table
Dim rsTemp As New ADODB.Recordset       'temporary recordset
Dim conn As New ADODB.Connection        'connection to small lakes
Dim cmd As New ADODB.Command

'declare variables for setting foriegn keys
Dim a_id As Variant                     'assessment id
Dim fc_id As Variant                    'fish collection id
Dim cs_id As Variant                    'creel survey id
Dim sp_id As Variant                    'species id
Dim cl_id As Variant                    'lookup clip id
Dim ss_id As Variant                    'lookup strain species id
Dim mt_id As Variant                    'lookup maturity id
Dim m_id As Variant                     'lookup method id
Dim ms_id As Variant                    'lookup mesh size id
Dim sc_id As Variant                    'lookup stomach contents id
Dim pl_id As Variant                    'lookup ploidy id
Dim am_id As Variant                    'age method id

Dim party As Variant                    'party_no variable
Dim person As Variant                   'person_no_variable

'open logfile
outfilepath = "\\FFSCOFP04\Users$\bonnie.robert\My Documents\Projects\DB Restructuring\Logfiles\biological_data_load_log.txt"
biold_log = FreeFile()
Close #biold_log
Open outfilepath For Output As #biold_log
Print #biold_log, "Starting QC error report: " & Now

'open connection to Small Lakes-Test
conn.Open "SMALL_LAKES-TEST", "GOFISHBC/Bonnie.Robert"

'open local Biological Data table (read only)
Set db = CurrentDb
Set rsData = db.OpenRecordset("SELECT * FROM ffsbc_Biological_Data;")

'open SLD individual_fish table to add new records
rsIF.CursorType = adOpenKeyset
rsIF.LockType = adLockOptimistic
rsIF.Open "SELECT * FROM ffsbc.individual_fish;", conn

'open Creel Survey records to add new records
rsCS.CursorType = adOpenKeyset
rsCS.LockType = adLockOptimistic
rsCS.Open "SELECT * FROM ffsbc.creel_survey;", conn

'open SLD project_assessment table table to add new records
rsProj_Asmnt.CursorType = adOpenKeyset
rsProj_Asmnt.LockType = adLockOptimistic
rsProj_Asmnt.Open "SELECT * FROM ffsbc.project_assessment;", conn

'open SLD assessment table to new add records
rsAsmnt.CursorType = adOpenKeyset
rsAsmnt.LockType = adLockOptimistic
rsAsmnt.Open "SELECT * FROM ffsbc.assessment;", conn

'for each record in Biological Data
Do Until rsData.EOF
    Debug.Print ("Processing Bio_Data_ID: " & rsData.Fields("Bio_Data_ID") & ": Row ID: " & rsData.AbsolutePosition + 1)
    
    'check to see if there's a related assessment in SLD
    a_key = rsData.Fields("Assessment_ID")
    
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.assessment WHERE assessment_key = '" & a_key & "';", conn
    
    'grab assessment_id
    If (rsTemp.RecordCount = 0) Then
        a_id = 0
        Print #biold_log, "Assessment " & a_key & " not listed in database."
    Else:
        a_id = rsTemp.Fields("assessment_id")
    End If
    rsTemp.Close
  
    'check to see if there's a related project in SLD
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT project_id FROM ffsbc.project WHERE project_name = '" & Trim(rsData.Fields("FFSBC_ID")) & "';", conn
                
    'grab project_id
    If (rsTemp.RecordCount = 0) Then
        p_id = 88
    Else:
        p_id = rsTemp.Fields("project_id")
    End If
    rsTemp.Close
  
    'open species to get related species id  for target species
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.species WHERE species_code = '" & Trim(rsData.Fields("Species")) & "';", conn
    
    If (rsTemp.RecordCount = 0) Then
        sp_id = Null
    Else:
        sp_id = rsTemp.Fields("species_id")
    End If
    rsTemp.Close

    'open method lookup table to get related lookup id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.lookup_method WHERE method_code = '" & Trim(rsData.Fields("Method")) & "';", conn
    
    If (rsTemp.RecordCount = 0) Then
        m_id = Null
    Else:
        m_id = rsTemp.Fields("lookup_method_id")
    End If

    rsTemp.Close
                                    
    'open mesh size lookup table to get related lookup id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.lookup_mesh_size WHERE mesh_size = CAST('" & Trim(rsData.Fields("Net_Mesh_Size")) & "' AS Float);", conn

    If (rsTemp.RecordCount = 0) Then
        ms_id = Null
    Else:
        ms_id = rsTemp.Fields("lookup_mesh_size_id")
    End If
        
    rsTemp.Close
                                    
    'open stomach_contents lookup table to get related lookup id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.lookup_stomach_content WHERE stomach_content_code = '" & Trim(rsData.Fields("Stomach_Contents")) & "';", conn
    
    If (rsTemp.RecordCount = 0) Then
        sc_id = Null
    Else:
        sc_id = rsTemp.Fields("lookup_stomach_content_id")
    End If

    rsTemp.Close
        
    'open age_method lookup table to get related lookup id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.lookup_age_method WHERE age_method = '" & Trim(rsData.Fields("Aged_Method")) & "';", conn
    
    If (rsTemp.RecordCount = 0) Then
        am_id = Null
    Else:
        am_id = rsTemp.Fields("lookup_age_method_id")
    End If
    
    rsTemp.Close
    
    'open maturity lookup table to get related lookup id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.lookup_maturity WHERE maturity_code = '" & Trim(rsData.Fields("Maturity")) & "';", conn

    If (rsTemp.RecordCount = 0) Then
        mt_id = Null
    Else:
        mt_id = rsTemp.Fields("lookup_maturity_id")
    End If

    rsTemp.Close
                                    
    'open strain_species lookup table to get related lookup id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.lookup_strain_species WHERE strain_species_code = '" & Trim(rsData.Fields("Strain")) & "';", conn
    
    If (rsTemp.RecordCount = 0) Then
        ss_id = Null
    Else:
        ss_id = rsTemp.Fields("lookup_strain_species_id")
    End If

    rsTemp.Close
    
    'open clip lookup table to get related lookup id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.lookup_clip WHERE clip_code = '" & Trim(rsData.Fields("Clip")) & "';", conn

    If (rsTemp.RecordCount = 0) Then
        cl_id = Null
    Else:
        cl_id = rsTemp.Fields("lookup_clip_id")
    End If

    rsTemp.Close

    
    'open ploidy lookup table to get related lookup id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.lookup_ploidy WHERE ploidy = '" & Trim(rsData.Fields("Ploidy")) & "';", conn

    If (rsTemp.RecordCount = 0) Then
        pl_id = Null
    Else:
        pl_id = rsTemp.Fields("lookup_ploidy_id")
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
                !end_date = Null
                !lookup_method_id = 4
                !date_added = rsData.Fields("Date_Added")
                !comments = rsData.Fields("Comments")
                !date_updated = Date
            .Update
            Print #biold_log, "Adding new assessment for " & a_key
            'a_id = 0
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
    
    'if method is 'CR' obtain creel survey id.  Throw an error if none exists
    If (m_id = 4) Then
    
        If IsNull(rsData.Fields("Party_No")) Then
            party = "Null"
        Else: party = rsData.Fields("Party_No")
        End If
        
        If IsNull(rsData.Fields("Person_No")) Then
            person = "Null"
        Else: person = rsData.Fields("Person_No")
        End If
        
        'open creel survey table to get related region id
        rsTemp.CursorType = adOpenKeyset
        rsTemp.LockType = adLockOptimistic
        rsTemp.Open "SELECT creel_survey_id FROM ffsbc.creel_survey WHERE assessment_id = " & a_id & " AND date_of_interview = CONVERT(datetime,'" _
                     & rsData.Fields("Date") & "', 104) AND party_no = " & party & " AND person_no = " & person & ";", conn
        
        'if no creel survey for this record, add a new creel survey record
        If (rsTemp.RecordCount = 0) Then
            
            With rsCS
                .AddNew
                !assessment_id = a_id
                !lookup_method_id = 4
                !date_of_interview = rsData.Fields("Date")
                !launch_or_fishing_location = rsData.Fields("Name")
                !party_no = rsData.Fields("Party_No")
                !Person_No = rsData.Fields("Person_No")
                !date_added = rsData.Fields("Date_Added")
                !comments = rsData.Fields("Comments")
                .Update
            End With
            Print #biold_log, "Adding new creel_survey for assessment" & a_key
            cs_id = rsCS.Fields("creel_survey_id")
            'cs_id = 0
            
        ElseIf (rsTemp.RecordCount > 1) Then
            Print #biold_log, "Adding new assessment for " & a_key
        Else
            cs_id = rsTemp.Fields("creel_survey_id")
        End If
        
        rsTemp.Close
        
    End If

    'create new record in biological_data_survey and enter data for current record
    With rsIF
            .AddNew
            !creel_survey_id = cs_id
            !assessment_id = a_id
            '!fish_collection_id = fc_id
            !species_id = sp_id
            !Date = rs.Data.Fields("Date")
            !fish_seq_no = rsData.Fields("Fish_ID")
            !lookup_clip_id = Null
            !clip = rsData.Fields("Clip")
            !lookup_strain_species_id = ss_id
            !length_mm = rsData.Fields("Length_mm")
            !weight_g = rsData.Fields("Weight_g")
            !sex = rsData.Fields("Sex")
            !age = rsData.Fields("Age")
            !lookup_maturity_id = mt_id
            !life_stage = rsData.Fields("Life_Stage")
            !Scale = rsData.Fields("Scale")
            !lookup_method_id = m_id
            !net_ID = rsData.Fields("Net_ID")
            !lookup_mesh_size_id = ms_id
            !otolith = rsData.Fields("Otolith")
            !atus = rsData.Fields("ATUS")
            '!lookup_stomach_content_id = sc_id
            !stomach_content = rsData.Fields("Stomach_Contents")
            !lookup_ploidy_id = pl_id
            !family_group = rsData.Fields("Family_Group")
            !tag_ID = rsData.Fields("Tag_ID")
            !lookup_age_method_id = am_id
            !comments = rsData.Fields("Comments")
            !date_added = rsData.Fields("Date_Added")
            .Update
    End With
    
'*****************************finished processing current data record
'set current record in biological_data data to next record
rsData.MoveNext
Loop

'close handles to commit changes
rsData.Close
rsCS.Close
rsIF.Close
rsAsmnt.Close
rsAsmnt_Proj.Close
conn.Close

Exit_load_biological_data:
    DoCmd.SetWarnings True
    Exit Sub

load_biological_data_Err:
    MsgBox Err.Description
    Resume Exit_load_biological_data
End Sub





