Attribute VB_Name = "load_effort_6"
Option Compare Database
Private Sub load_effort()
'loads Creel data into creel_survey
'includes a corresponding entry in assessment_id
'links to survey_type_id
'populates creel_fish_counts

On Error GoTo load_effort_Err

Dim db As DAO.Database
Dim rsData As DAO.Recordset             'data to load
Dim rsEf As New ADODB.Recordset         'effort table
Dim rsAsmnt As New ADODB.Recordset      'assessment table
Dim rsProj_Asmnt As New ADODB.Recordset 'project-assessment table
Dim rsTemp As New ADODB.Recordset       'temporary recordset
Dim conn As New ADODB.Connection        'connection to small lakes

Dim cmd As New ADODB.Command
Dim match As Boolean

Dim a_id As Variant
Dim cs_id As Variant
Dim s As String
Dim sp_code As String
Dim sp_id As Variant
Dim target_sp_id As Variant

outfilepath = "\\FFSCOFP04\Users$\bonnie.robert\My Documents\Projects\DB Restructuring\Logfiles\effort_load_log.txt"
efld_log = FreeFile()
Close #efld_log
Open outfilepath For Output As #efld_log
Print #efld_log, "Starting QC error report: " & Now

'open connection to small lakes test database
conn.Open "SMALL_LAKES-TEST", "GOFISHBC/Bonnie.Robert"

'open Creel table in local Access instance
Set db = CurrentDb
Set rsData = db.OpenRecordset("SELECT * FROM ffsbc_Effort_Counts;")

'open creel_survey table to add new records
rsEf.CursorType = adOpenKeyset
rsEf.LockType = adLockOptimistic
rsEf.Open "SELECT * FROM ffsbc.fishing_effort;", conn

'open project_assessment table table to add new records
rsProj_Asmnt.CursorType = adOpenKeyset
rsProj_Asmnt.LockType = adLockOptimistic
rsProj_Asmnt.Open "SELECT * FROM ffsbc.project_assessment;", conn

'open assessment table to new add records
rsAsmnt.CursorType = adOpenKeyset
rsAsmnt.LockType = adLockOptimistic
rsAsmnt.Open "SELECT * FROM ffsbc.assessment;", conn

'for each record in effort dataset (read only)
Do Until rsData.EOF
    Debug.Print ("Processing Effort_ID: " & rsData.Fields("Effort_ID"))
    
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

    'open waterbody table to get primary key id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.waterbody WHERE MOF_waterbody_id = '" & Trim(rsData.Fields("WBID")) & "';", conn
                      
    If (rsTemp.RecordCount = 0) Then
        wb_id = 0
    Else:
        wb_id = rsTemp.Fields("waterbody_id")
    End If
    rsTemp.Close
    
    'open region table to get primary key id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.waterbody_region WHERE waterbody_id = '" & wb_id & "';", conn
                      
    If (rsTemp.RecordCount = 0) Then
        r_id = 13
    Else:
        r_id = rsTemp.Fields("region_id")
    End If
    rsTemp.Close
  
    'open lookup_method to get primary key id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.lookup_method WHERE method_code = '" & Trim(rsData.Fields("Method")) & "';", conn
    
    If (rsTemp.RecordCount = 0) Then
        m_id = 23
    Else:
        m_id = rsTemp.Fields("lookup_method_id")
    End If
    rsTemp.Close

    'open lookup_weather to get primary key id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.lookup_weather WHERE weather = '" & Trim(rsData.Fields("Weather")) & "';", conn
    
    If (rsTemp.RecordCount = 0) Then
        we_id = Null
    Else:
        we_id = rsTemp.Fields("lookup_weather_id")
    End If
    rsTemp.Close

    'open lake_condition to get primary key id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.lookup_lake_condition WHERE lake_condition = '" & Trim(rsData.Fields("Lake_Condition")) & "';", conn
    
    If (rsTemp.RecordCount = 0) Then
        lc_id = Null
    Else:
        lc_id = rsTemp.Fields("lookup_lake_condition_id")
    End If
    rsTemp.Close
  
    'open image_quality to get primary key id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.lookup_image_quality WHERE image_quality = '" & Trim(rsData.Fields("Image_Quality")) & "';", conn
    
    If (rsTemp.RecordCount = 0) Then
        iq_id = Null
    Else:
        iq_id = rsTemp.Fields("lookup_image_quality_id")
    End If
    rsTemp.Close

    'open camera_type to get primary key id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.lookup_camera_type WHERE camera_type = '" & Trim(rsData.Fields("Camera_Type")) & "';", conn
    
    If (rsTemp.RecordCount = 0) Then
        c_id = Null
    Else:
        c_id = rsTemp.Fields("lookup_camera_type_id")
    End If
    rsTemp.Close

    'add assessment and assessment/project record if none in assessment table
    If (a_id = 0) Then
        With rsAsmnt
            .AddNew
                !waterbody_id = wb_id
                !Assessment_Key = a_key
                !region_id = r_id
                !start_date = rsData.Fields("Date")
                !end_date = rsData.Fields("Date")
                !lookup_method_id = m_id
                !date_added = rsData.Fields("Date_Added")
                !comments = rsData.Fields("Comments") & " data transferred from old SLD database."
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
    With rsEf
            .AddNew
            !assessment_id = a_id
            !Date = rsData.Fields("Date")
            !Time = Left(rsData.Fields("Time"), 8)
            !camera_or_lake_location = rsData.Fields("Name")
            !lookup_method_id = m_id
            !lookup_weather_id = we_id
            !lookup_lake_condition_id = lc_id
            !boats_1angler = rsData.Fields("Boats_1Angler")
            !boats_2anglers = rsData.Fields("Boats_2Anglers")
            !boats_3anglers = rsData.Fields("Boats_3Anglers")
            !boats_4anglers = rsData.Fields("Boats_4Anglers")
            !boats_5anglers = rsData.Fields("Boats_5Anglers")
            !boats_unknown = rsData.Fields("Boats_Unknown")
            !boats_not_fishing = rsData.Fields("Boats_NOT_Fishing")
            !shore_ice_anglers = rsData.Fields("Shore_Ice_Anglers")
            !shore_ice_not_fishing = rsData.Fields("Shore_Ice_NOT_Fishing")
            !shore_ice_unknown = rsData.Fields("Shore_Ice_Unknown")
            !ice_fishing_tents_1angler = rsData.Fields("Ice_Fishing_Tents_1Angler")
            !ice_fishing_tents_2anglers = rsData.Fields("Ice_Fishing_Tents_2Anglers")
            !ice_fishing_tents_unknown = rsData.Fields("Ice_Fishing_Tents_Unknown")
            !lookup_image_quality_id = iq_id
            !lookup_camera_type_id = c_id
            !percent_visibility = rsData.Fields("Percent_Visibility")
            !portion_of_lake_seen = rsData.Fields("Portion_of_Lake_Seen")
            !picture_name = rsData.Fields("Picture_Name")
            !folder_name = rsData.Fields("Folder_Name")
            !sampler = rsData.Fields("Sampler")
            !page_ID = rsData.Fields("Page_ID")
            !time_arrive = rsData.Fields("Time_Arrive")
            !time_leave = rsData.Fields("Time_Leave")
            !date_added = rsData.Fields("Date_Added")
            !comments = rsData.Fields("Comments")
            .Update
    End With

'*****************************finished processing current data record
'set current record in Creel data to next record
rsData.MoveNext
Loop

'close handles to commit changes
rsData.Close
rsCS.Close
rsFC.Close
rsAsmnt.Close
rsAsmnt_Proj.Close
conn.Close

Exit_load_effort:
    DoCmd.SetWarnings True
    Exit Sub

load_effort_Err:
    MsgBox Err.Description
    Resume Exit_load_effort
End Sub



