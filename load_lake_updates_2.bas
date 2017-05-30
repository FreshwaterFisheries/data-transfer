Attribute VB_Name = "load_lake_updates_2"
Private Sub load_lake_updates()
'loads lake update data into waterbody dimensions table
'assumes qc on entries has been completed and there is a match in lookup tables.
'assumes dates are valid, parameters are valid
'errors are logged to: ld_lkup_log

On Error GoTo load_lake_updates_Err

Dim db As DAO.Database
Dim rsData As DAO.Recordset             'data to load

Dim rsWb As New ADODB.Recordset         'waterbody table
Dim rsAsmnt As New ADODB.Recordset      'assessment table
Dim rsProj_Asmnt As New ADODB.Recordset 'project_assessment table
Dim rsWbd As New ADODB.Recordset        'waterbody dimensions recordset
Dim rsTemp As New ADODB.Recordset       'miscellaneous recordset

Dim conn As New ADODB.Connection        'connection to small lakes
Dim cmd As New ADODB.Command

Dim wb_id As Variant                    'waterbody id
Dim a_id As Variant                     'assessment id
Dim a_key As Variant                    'assessment_key

Dim m_id As Variant                     'lookup_method id
Dim r_id As Variant                     'region id

Dim null_proj_id As Integer             'null project id in project table
Dim null_m_id As Integer                'null method id in project table
Dim m_name As Variant                   'method code
Dim dat As Variant                      'date string for assessment key

Dim outfilepath As String               'filename for output

null_m_id = 26
null_proj_id = 88

outfilepath = "\\FFSCOFP04\Users$\bonnie.robert\My Documents\Projects\DB Restructuring\Logfiles\lake_update_ld_log.txt"
ld_lkup_log = FreeFile()
Close #ld_lkup_log
Open outfilepath For Output As #ld_lkup_log
Print #ld_lkup_log, "Starting ld error report: " & Now

'open connection to small lakes test database
conn.Open "SMALL_LAKES-TEST", "GOFISHBC/Bonnie.Robert"

'open Sampling Summary table in local Access instance
Set db = CurrentDb
Set rsData = db.OpenRecordset("SELECT * FROM ffsbc_Lake_Size_Updates;")

'open project_assessment table table to add new records
rsProj_Asmnt.CursorType = adOpenKeyset
rsProj_Asmnt.LockType = adLockOptimistic
rsProj_Asmnt.Open "SELECT * FROM ffsbc.project_assessment;", conn

'open assessment table to new add records
rsAsmnt.CursorType = adOpenKeyset
rsAsmnt.LockType = adLockOptimistic
rsAsmnt.Open "SELECT * FROM ffsbc.assessment;", conn

'open assessment table to new add records
rsWbd.CursorType = adOpenKeyset
rsWbd.LockType = adLockOptimistic
rsWbd.Open "SELECT * FROM ffsbc.waterbody_dimensions;", conn

'for each record in Lakes_Profiles dataset (read only)
Do Until rsData.EOF
    
    Debug.Print ("Loading Lake_Size_Update_ID: " & rsData.Fields("Lake_Size_Update_ID"))
    '******* get foreign key id's: method_id,waterbody_id and (if exists) assessment_id
                         
    'open method lookup table to get related lookup method id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.lookup_method WHERE method_code = '" & Trim(rsData.Fields("Method")) & "';", conn
    
    If (rsTemp.RecordCount = 0) Then
        m_id = null_m_id
        m_name = "UK"
    Else:
        m_id = rsTemp.Fields("lookup_method_id")
        m_name = rsTemp.Fields("method_code")
    End If
    rsTemp.Close
                                    
    'open waterbody table to get related waterbody id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.waterbody WHERE MOF_waterbody_id = '" & Trim(rsData.Fields("WBID")) & "';", conn
                      
    If (rsTemp.RecordCount = 0) Then
        'Print #ld_lkup_log, "No waterbody associated with assessment " & rsData.Fields("Assessment_ID")
    Else:
        wb_id = rsTemp.Fields("waterbody_id")
    End If
    rsTemp.Close
    
    'open region table to get related region id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.waterbody_region WHERE waterbody_id = '" & wb_id & "';", conn
                      
    If (rsTemp.RecordCount = 0) Then
        Print #ld_lkup_log, "No region associated with lake_size_update ID " & rsData.Fields("Lake_Size_Update_ID")
    Else:
        r_id = rsTemp.Fields("region_id")
    End If
    rsTemp.Close
    
    If (IsNull(rsData.Fields("Date"))) Then
            dat = "00000"
    Else:   dat = CStr(CLng(rsData.Fields("Date")))
    End If
    
    'open assessment table to see if there's an associated assessment
    a_key = rsData.Fields("WBID") & "_" & dat & "_" & m_name
    
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.assessment WHERE assessment_key = '" & a_key & "';", conn
    
    If (rsTemp.RecordCount = 0) Then
        a_id = 0
    Else:
        a_id = rsTemp.Fields("assessment_id")
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
                !lookup_method_id = m_id
                !date_added = Date
                !comments = rsData.Fields("Comments") & "Updates from old SL database table 1c - Lake_Size_Updates"
                !date_updated = Date
            .Update
            a_id = rsAsmnt.Fields("assessment_id")
        End With
        
        'add assessment_id and project_id to assessment_project table
        With rsProj_Asmnt
            .AddNew
                !project_id = null_proj_id
                !assessment_id = a_id
            .Update
        End With
        
     End If
   
    'create new record in waterbody_dimensions and enter data for current record
    With rsWbd
        .AddNew
        !assessment_id = a_id
        !waterbody_id = wb_id
        !area_surface_ha = rsData.Fields("Area_Surface_ha")
        !lake_volume_m3 = rsData.Fields("Lake_Volume_m3")
        !depth_max_m = rsData.Fields("Depth_Max_m")
        !depth_mean_m = rsData.Fields("Depth_Mean_m")
        !source_lake_parameters = rsData.Fields("Source")
        !date_lake_parameters = rsData.Fields("Date")
        !date_added = Date
        !comments = rsData.Fields("Comments") & ". Updates from old SL database table 1c - Lake_Size_Updates"
        .Update
    End With

    '*****************************finished processing current data record
    'set current record in lake_profile data to next record
    rsData.MoveNext
Loop

'close handles to commit changes
rsData.Close
rsWbd.Close
rsAsmnt.Close
rsProj_Asmnt.Close
conn.Close

Exit_load_lake_updates:
    DoCmd.SetWarnings True
    Exit Sub

load_lake_updates_Err:
    MsgBox Err.Description
    Resume Exit_load_lake_updates
End Sub
