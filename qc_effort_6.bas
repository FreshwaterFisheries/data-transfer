Attribute VB_Name = "qc_effort_6"
Option Compare Database

Private Sub qc_effort()
'loads sampling_summary data into waterbody profile and measurement tables
'includes a corresponding entry in assessment_id ***This part needs to be addressed!

On Error GoTo qc_effort_Err

Dim db As DAO.Database
Dim rsData As DAO.Recordset             'data to load
Dim rsAsmnt As New ADODB.Recordset      'assessment table
Dim rsTemp As New ADODB.Recordset       'waterbody table
Dim rsProj_Asmnt As New ADODB.Recordset 'project_assessment join table
Dim conn As New ADODB.Connection        'connection to small lakes
Dim cmd As New ADODB.Command

Dim p_id As Integer                      'project id
Dim a_id As Integer                      'assessment id

Dim wb_id As Variant                     'waterbody id
Dim r_id As Integer                      'region id

Dim m_id As Integer                      'method id
Dim we_id As Integer                     'weather_id
Dim lc_id As Integer                     'lake_condition_id
Dim lq_id As Integer                     'image_quality_id
Dim c_id As Integer                      'camera_type

'open connection to small lakes test database
conn.Open "SMALL_LAKES-TEST", "GOFISHBC/Bonnie.Robert"

'open Sampling Summary table in local Access instance
Set db = CurrentDb
Set rsData = db.OpenRecordset("SELECT * FROM ffsbc_Effort_Counts;")

outfilepath = "\\FFSCOFP04\Users$\bonnie.robert\My Documents\Projects\DB Restructuring\Logfiles\effort_qc_log.txt"
qc_effort_log = FreeFile()
Close #qc_effort_log
Open outfilepath For Output As #qc_effort_log
Print #qc_effort_log, "Starting QC error report: " & Now

'for each record in Effort Counts dataset (read only)
Do Until rsData.EOF
    Debug.Print "Procesing Record: " & rsData.Fields("Effort_ID")

    'open project table to get related project id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.project WHERE project_name = '" & Trim(rsData.Fields("FFSBC_ID")) & "';", conn
    If (rsTemp.RecordCount = 0) Then
        p_id = 88
    Else: p_id = rsTemp.Fields("project_id")
    End If
    rsTemp.Close
            
    'open waterbody table to get related waterbody id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.waterbody WHERE MOF_waterbody_id = '" & Trim(rsData.Fields("WBID")) & "';", conn
    If (rsTemp.RecordCount = 0) Then
        Print #qc_effort_log, "Missing waterbody for Effort_ID: " & rsData.Fields("Effort_ID")
        wb_id = 0
        r_id = 13
    Else: wb_id = rsTemp.Fields("waterbody_id")
    End If
    rsTemp.Close
                
    'open waterbody_region table to get related lookup_method id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.waterbody_region WHERE waterbody_id = '" & wb_id & "';", conn
    If (rsTemp.RecordCount = 0) Then
        Print #qc_effort_log, "Missing region for Effort_ID: " & rsData.Fields("Effort_ID")
        wb_id = 0
        r_id = 13
    Else: r_id = rsTemp.Fields("region_id")
    End If
    rsTemp.Close
                
    'open lookup_method table to get related lookup_method id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.lookup_method WHERE method_code = '" & Trim(rsData.Fields("Method")) & "';", conn
    If (rsTemp.RecordCount = 0) Then
        Print #qc_effort_log, "Missing method for Effort_ID: " & rsData.Fields("Effort_ID")
        m_id = 0
    Else: m_id = rsTemp.Fields("lookup_method_id")
    End If
    rsTemp.Close
    
    'open lookup_weather to get related lookup_weather id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.lookup_weather WHERE weather = '" & Trim(rsData.Fields("Weather")) & "';", conn
    If (rsTemp.RecordCount = 0) Then
        'Print #qc_effort_log, "Missing weather for Effort_ID: " & rsData.Fields("Effort_ID")
    Else: w_id = rsTemp.Fields("lookup_weather_id")
    End If
    rsTemp.Close
    
    'open lookup_lake_condition to get related lookup_lake_condition id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.lookup_lake_condition WHERE lake_condition = '" & Trim(rsData.Fields("Lake_Condition")) & "';", conn
    If (rsTemp.RecordCount = 0) Then
        'Print #qc_effort_log, "Missing lake_condition for Effort_ID: " & rsData.Fields("Effort_ID")
    Else: l_id = rsTemp.Fields("lookup_lake_condition_id")
    End If
    rsTemp.Close
    
    'open assessment table to get assessment id
    rsAsmnt.CursorType = adOpenKeyset
    rsAsmnt.LockType = adLockOptimistic
    rsAsmnt.Open "SELECT * FROM ffsbc.assessment WHERE assessment_key = '" & Trim(rsData.Fields("Assessment_ID")) & "';", conn
    'assessment_id found in assessment table
    Select Case rsAsmnt.RecordCount
        
        'more than one, throw error
        Case Is > 1
            Print #N, ("multiple assessment_id's found for ffsbc.assessment: " & rsData.Fields("Assessment_ID"))
    
        'no assessment id, do nothing
        Case Is = 0
            
        Case Is = 1
            'should never happen
            If IsNull(rsAsmnt.Fields("assessment_id")) Then
                Print #N, "Null Assessment_ID in record #" & rsData.AbsolutePosition
            'otherwise, get the foriegn key id's for assessment, project, region, method, and waterbody that
            'are associated with the assessment table
            Else: a_id = (rsAsmnt.Fields("assessment_id"))
            End If
            
            'check for mismatched WBID
            If Not (wb_id = rsAsmnt.Fields("waterbody_id")) Then
                Print #qc_effort_log, "error in record " & rsData.Fields(0) & _
                "     assessment_id:" & rsData.Fields("Assessment_id") & _
                "     uploading WBID:" & wb_id & _
                "     database waterbody_id:" & rsAsmnt.Fields("waterbody_id")
            End If
    
            'check for mismatched method
            If Not (m_id = rsAsmnt.Fields("lookup_method_id")) Then
                Print #qc_effort_log, "error in record " & rsData.Fields(0) & _
                "     assessment_key:" & rsData.Fields("Assessment_id") & _
                "     uploading Method:" & rsData.Fields("Method") & _
                "     database lookup_method_id:" & rsAsmnt.Fields("lookup_method_id")
            End If
        
            'check for mismatched project_ID
            rsProj_Asmnt.CursorType = adOpenKeyset
            rsProj_Asmnt.LockType = adLockOptimistic
            rsProj_Asmnt.Open "SELECT * FROM ffsbc.project_assessment WHERE project_id = " & p_id & " AND assessment_id = " & a_id & ";", conn
            
            If rsProj_Asmnt.RecordCount = 0 Then
              Print #qc_effort_log, "error in record " & rsData.Fields(0) & _
              "     assessment_key: " & rsData.Fields("Assessment_id") & _
              " is not related to project: " & rsData.Fields("FFSBC_ID") & _
              " in the database"
              'Add project assessment
            End If
        
            rsProj_Asmnt.Close
        
    End Select
    
    'close handles to commit changes
    rsAsmnt.Close
    

    '*****************************finished processing current data record
    'set current record in Sampling Summary data to next record
    rsData.MoveNext
Loop

'close handles to commit changes
rsData.Close
Close #f
conn.Close

Exit_qc_effort:
    DoCmd.SetWarnings True
    Exit Sub

qc_effort_Err:
    MsgBox Err.Description
    Resume Exit_qc_effort
End Sub

