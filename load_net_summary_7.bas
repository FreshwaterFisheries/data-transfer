Attribute VB_Name = "load_net_summary_7"
Option Compare Database

Private Sub load_net_summary()
'loads Creel data into creel_survey
'includes a corresponding entry in assessment_id
'links to survey_type_id
'populates creel_fish_counts

On Error GoTo load_net_summary_Err

Dim db As DAO.Database
Dim rsData As DAO.Recordset             'data to load
Dim rsFC As New ADODB.Recordset         'fish_collection table
Dim rsFCC As New ADODB.Recordset        'fish_collection_count table
Dim rsAsmnt As New ADODB.Recordset      'assessment table
Dim rsProj_Asmnt As New ADODB.Recordset 'project-assessment table
Dim rsTemp As New ADODB.Recordset       'temporary recordset
Dim conn As New ADODB.Connection        'connection to small lakes

Dim a_id As Variant                     'assessment id
Dim p_id As Integer                     'project id
Dim wb_id As Variant                    'waterbody id
Dim m_id As Integer                     'method id
Dim r_id As Integer                     'region id
Dim fc_id As Integer                    'net_summary_survey id
Dim sp_id As Integer                    'species_id
Dim sd_id As Integer                    'sample design id
Dim s_id As Integer                     'setting id
Dim h_id As Integer                     'habitat_id
Dim src As String                       'source field

outfilepath = "\\FFSCOFP04\Users$\bonnie.robert\My Documents\Projects\DB Restructuring\Logfiles\net_summary_load_log.txt"
net_log = FreeFile()
Close #net_log
Open outfilepath For Output As #net_log
Print #net_log, "Starting QC error report: " & Now

'open connection to small lakes test database
conn.Open "SMALL_LAKES-TEST", "GOFISHBC/Bonnie.Robert"

'open Creel table in local Access instance
Set db = CurrentDb
Set rsData = db.OpenRecordset("SELECT * FROM ffsbc_Net_Summary;")

'open fish_count table for adding new records
rsFC.CursorType = adOpenKeyset
rsFC.LockType = adLockOptimistic
rsFC.Open "SELECT * FROM ffsbc.fish_collection;", conn

'open creel_survey table to add new records
rsFCC.CursorType = adOpenKeyset
rsFCC.LockType = adLockOptimistic
rsFCC.Open "SELECT * FROM ffsbc.fish_collection_count;", conn

'open project_assessment table table to add new records
rsProj_Asmnt.CursorType = adOpenKeyset
rsProj_Asmnt.LockType = adLockOptimistic
rsProj_Asmnt.Open "SELECT * FROM ffsbc.project_assessment;", conn

'open assessment table to new add records
rsAsmnt.CursorType = adOpenKeyset
rsAsmnt.LockType = adLockOptimistic
rsAsmnt.Open "SELECT * FROM ffsbc.assessment;", conn

'for each record in net_summary dataset (read only)
rsData.Move (346)
Do Until rsData.EOF
    Debug.Print ("Processing Gillnet_Summary_ID: " & rsData.Fields("Gillnet_Summary_ID"))
    
    'open assessment table to see if there's an associated assessment
    a_key = rsData.Fields("Assessment_ID")
    
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.assessment WHERE assessment_key = '" & a_key & "';", conn
    
    If (rsTemp.RecordCount = 0) Then
        a_id = 0
    Else:
        a_id = rsTemp.Fields("assessment_id")
        src = rsTemp.Fields("source")
    End If
    rsTemp.Close
    
    'concatenate two sources if they are different
    If Not (src = rsData.Fields("Source")) Then
        
        If (Not (IsNull(src)) And Not (IsNull(rsData.Fields("Source")))) Then
            src = src & ";"
        End If
    
        src = src & rsData.Fields("Source")
    
    End If
    
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
  
    'open method to get related method id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.lookup_method WHERE method_code = '" & Trim(rsData.Fields("Method")) & "';", conn
    
    If (rsTemp.RecordCount = 0) Then
        m_id = 23
    Else:
        m_id = rsTemp.Fields("lookup_method_id")
    End If
    rsTemp.Close
  
    'open species to get related species id  for target species
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.species WHERE species_code = '" & Trim(rsData.Fields("Species_Code")) & "';", conn
    
    If (rsTemp.RecordCount = 0) Then
        sp_id = 62
    Else:
        sp_id = rsTemp.Fields("species_id")
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
    
    'open habitat table to get related habitat_type id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.lookup_habitat_type WHERE habitat_type_code = '" & Trim(rsData.Fields("Habitat")) & "';", conn
                      
    If (rsTemp.RecordCount = 0) Then
        h_id = Null
    Else:
        h_id = rsTemp.Fields("lookup_habitat_type_id")
    End If
    rsTemp.Close
    
    'open setting table to get related setting id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.lookup_setting WHERE setting_code = '" & Trim(rsData.Fields("Setting")) & "';", conn
                      
    If (rsTemp.RecordCount = 0) Then
        s_id = Null
    Else:
        s_id = rsTemp.Fields("lookup_setting_id")
    End If
    rsTemp.Close
    
    'open sample_design table to get related sample_design id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.lookup_sample_design WHERE sample_design_code = '" & Trim(rsData.Fields("Sample_Design")) & "';", conn
                      
    If (rsTemp.RecordCount = 0) Then
        sd_id = Null
    Else:
        sd_id = rsTemp.Fields("lookup_sample_design_id")
    End If
    rsTemp.Close

    'add assessment and assessment/project record if none in assessment table
    If (a_id = 0) Then
        With rsAsmnt
            .AddNew
                !waterbody_id = wb_id
                !Assessment_Key = a_key
                !region_id = r_id
                !Source = src
                !start_date = rsData.Fields("Start_Date")
                !end_date = rsData.Fields("End_Date")
                !lookup_method_id = m_id
                !date_added = Date
                !comments = "Net Summary data transferred from old SL database table 2a - Net Summary"
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
    With rsFC
            .AddNew
            !assessment_id = a_id
            !lookup_method_id = m_id
            !shore_utm_zone = rsData.Fields("Shore_Zone")
            !shore_utm_easting = rsData.Fields("Shore_Easting")
            !shore_utm_northing = rsData.Fields("Shore_Northing")
            !lake_utm_zone = rsData.Fields("Lake_Zone")
            !lake_utm_easting = rsData.Fields("Lake_Easting")
            !lake_utm_northing = rsData.Fields("Lake_Northing")
            !lookup_setting_id = s_id
            !lookup_habitat_id = h_id
            !start_date = rsData.Fields("Start_Date")
            !end_date = rsData.Fields("End_Date")
            !start_time = rsData.Fields("Start_Time")
            !end_time = rsData.Fields("End_Time")
            !haul_pass = rsData.Fields("Haul_Pass")
            !net_panels = rsData.Fields("Net_panels")
            !net_sinking_or_floating = rsData.Fields("Net_sinking_or_floating")
            !overnight_yn = rsData.Fields("Overnight_YN")
            !net_ID = rsData.Fields("Net_ID")
            !lookup_sample_design_id = sd_id
            !Source = src
            !site = rsData.Fields("Site")
            !report_authors = rsData.Fields("Report_Authors")
            !date_added = rsData.Fields("Date_Added")
            !comments = rsData.Fields("Comments")
            .Update
            'get primary key to enter in creel_fish_counts
            fc_id = rsFC.Fields("fish_collection_id")
            
    End With
    
    'create new record in creel_archived and enter archived data for current record
    With rsFCC
    
        .AddNew
        !fish_collection_id = fc_id
        !species_id = sp_id
        !Count = rsData.Fields("Count_Fish")
        .Update
    
    End With


'*****************************finished processing current data record
'set current record in Creel data to next record
rsData.MoveNext
Loop

'close handles to commit changes
rsData.Close
rsFCC.Close
rsFC.Close
rsAsmnt.Close
rsAsmnt_Proj.Close
conn.Close

Exit_load_net_summary:
    DoCmd.SetWarnings True
    Exit Sub

load_net_summary_Err:
    MsgBox Err.Description
    Resume Exit_load_net_summary
End Sub
