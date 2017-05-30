Attribute VB_Name = "qc_sampling_summary_1"
Option Compare Database

Private Sub qc_sampling_summary()
'qc's sampling_summary data for loading into waterbody profile and measurement tables
'includes a corresponding entry in assessment_id

On Error GoTo qc_sampling_summary_Err

Dim db As DAO.Database
Dim rsData As DAO.Recordset             'data to load
Dim rsAsmnt As New ADODB.Recordset      'assessment table
Dim rsProj As New ADODB.Recordset       'project table
Dim rsWb As New ADODB.Recordset         'waterbody table
Dim rsSp As New ADODB.Recordset         'historical catch summary table
Dim rsMd As New ADODB.Recordset         'lookup method table
Dim rsProj_Asmnt As New ADODB.Recordset 'project_assessment join table
Dim rsTemp As New ADODB.Recordset       'temporary recordset

Dim conn As New ADODB.Connection        'connection to small lakes
Dim cmd As New ADODB.Command

Dim match As Boolean
Dim cntr As Integer
Dim wb_name As String                    'waterbody name
Dim a_id As Integer                      'assessment id
Dim p_id As Integer                      'project id
Dim wb_id As Variant                     'waterbody id
Dim r_id As Integer                      'region id
Dim m_id As Integer                      'method id

qc_log = FreeFile()
Close #qc_log
Open "\\FFSCOFP04\Users$\bonnie.robert\My Documents\Projects\DB Restructuring\Logfiles\sampling_summary_qc_log.txt" For Append As #qc_log
Print #qc_log, "Starting QC error report: " & Date

'open connection to small lakes test database
conn.Open "SMALL_LAKES-TEST", "GOFISHBC/Bonnie.Robert"

'open Sampling Summary table in local Access instance
Set db = CurrentDb
Set rsData = db.OpenRecordset("SELECT * FROM ffsbc_Sampling_Summary;")

'for each record in Sampling Summary dataset (read only)
Do Until rsData.EOF
    Debug.Print "QC'ing Sampling_ID: " & rsData.Fields("Sampling_ID")
    
    '***********************************qc some values in Sampling Summary fields
    If (rsData.Fields("Start_Date") > rsData.Fields("End_Date")) Then
        Print #qc_log, "Date error. Sampling_ID " & rsData.Fields("Sampling_ID")
    End If
    
    If (rsData.Fields("Start_Date") > Date) Then
        Print #qc_log, "Date error. Sampling_ID " & rsData.Fields("Sampling_ID")
    End If
    
    If (rsData.Fields("End_Date") > Date) Then
        Print #qc_log, "Date error. Sampling_ID " & rsData.Fields("Sampling_ID")
    End If
    
    'pH qc
    If Not (4 < rsData.Fields("pH") < 10.5 Or IsNull(rsData.Fields("pH"))) Then
        Print #qc_log, "pH error. Row " & rsData.AbsolutePosition
    End If
    
    '**********************************check for NULL lookups, set to unknown
    If IsNull(rsData.Fields("FFSBC_ID")) Then
        With rsData
            .Edit
            !FFSBC_ID = "UNK"
            .Update
        End With
    End If
    
    If IsNull(rsData.Fields("Method")) Then
        With rsData
            .Edit
            !method = "UNK"
            .Update
        End With
    End If
        
    If IsNull(rsData.Fields("Species_Code")) Then
        With rsData
            .Edit
            !method = "UNK"
            .Update
        End With
    End If
    
    '********************************************obtain recordsets from lookup, assessment and project tables
    'open assessment table to get assessment id
    rsAsmnt.CursorType = adOpenKeyset
    rsAsmnt.LockType = adLockOptimistic
    rsAsmnt.Open "SELECT * FROM ffsbc.assessment WHERE assessment_key = '" & Trim(rsData.Fields("Assessment_ID")) & "';", conn
    
    'open project table to get related project id
    rsProj.CursorType = adOpenKeyset
    rsProj.LockType = adLockOptimistic
    rsProj.Open "SELECT * FROM ffsbc.project WHERE project_name = '" & Trim(rsData.Fields("FFSBC_ID")) & "';", conn
            
    'open waterbody table to get related waterbody id
    rsWb.CursorType = adOpenKeyset
    rsWb.LockType = adLockOptimistic
    rsWb.Open "SELECT * FROM ffsbc.waterbody WHERE MOF_waterbody_id = '" & Trim(rsData.Fields("WBID")) & "';", conn
                
    'open method lookup table to get related lookup method id
    rsMd.CursorType = adOpenKeyset
    rsMd.LockType = adLockOptimistic
    rsMd.Open "SELECT * FROM ffsbc.lookup_method WHERE method_code = '" & Trim(rsData.Fields("Method")) & "';", conn
    
    'open species to get related species id
    rsSp.CursorType = adOpenKeyset
    rsSp.LockType = adLockOptimistic
    rsSp.Open "SELECT * FROM ffsbc.species WHERE species_code = '" & Trim(rsData.Fields("Species_Code")) & "';", conn
    
    '*********************************************check for valid matches in lookup and project tables
    'check for valid project in project table
    Select Case rsProj.RecordCount
        'if no related project, throw an error
        Case Is = 0
            Print #qc_log, "Missing project error.  Sampling_ID " & rsData.Fields("Sampling_ID") & ". Project " & rsData.Fields("FFSBC_ID")
              
        'if multiple related project, throw an error
        Case Is > 1
            Print #qc_log, "Multiple project error.  Sampling_ID " & rsData.Fields("Sampling_ID") & ". Project " & rsData.Fields("FFSBC_ID")

    End Select
                                        
    Select Case rsWb.RecordCount
        'if no related waterbody, throw an error
        Case Is = 0
            Print #qc_log, "Missing waterbody error.  Sampling_ID " & rsData.Fields("Sampling_ID") & ". WBID " & rsData.Fields("WBID")
        'if multiple related project, throw an error
        Case Is > 1
            Print #qc_log, "Multiple waterbodies error.  Sampling_ID " & rsData.Fields("Sampling_ID") & ". WBID " & rsData.Fields("WBID")
            
        'check that the waterbody name matches the waterbody name in the database
        Case Is = 1
            If Not (IsNull(rsData.Fields("Name"))) Then
                wb_name = rsData.Fields("Name")
                wb_name = Replace(wb_name, " Lake", "")
                wb_name = Replace(wb_name, " Creek", "")
        
                If (InStr(1, LCase(rsWb.Fields("gazetted_name")), LCase(wb_name)) = 0) Then
                    If (InStr(1, LCase(rsWb.Fields("alias")), LCase(wb_name)) = 0) Then
                        Print #qc_log, "Waterbody name error. Sampling_ID " & rsData.Fields("Sampling_ID") & " WBID " & rsData.Fields("WBID")
                    End If
                End If
            End If

    End Select
                                      
    Select Case rsMd.RecordCount
        'if no related method, throw an error
        Case Is = 0
            Print #qc_log, "Missing methods error.  Sampling_ID " & rsData.Fields("Sampling_ID") & ". Method " & rsData.Fields("Method")
            'if multiple related method, throw an error
        Case Is > 1
            Print #qc_log, "Multiple methods error.  Sampling_ID " & rsData.Fields("Sampling_ID") & ". Method " & rsData.Fields("Method")
    End Select
    
    Select Case rsSp.RecordCount
        'if no related species, throw an error
        Case Is = 0
            'Print #qc_log, "Missing species error.  Sampling_ID " & rsData.Fields("Sampling_ID") & ". Species " & rsData.Fields("Species_Code")
            'if multiple related species, throw an error
        Case Is > 1
            Print #qc_log, "Multiple species error.  Sampling_ID " & rsData.Fields("Sampling_ID") & ". Species " & rsData.Fields("Species_Code")
    End Select
    
    '****************************************************check values in matching entry in assessment table, if it exists
    'assessment_id found in assessment table
    Select Case rsAsmnt.RecordCount
        
        'more than one, throw error
        Case Is > 1
            Print #qc_log, "Multiple assessments error.  Sampling_ID " & rsData.Fields("Sampling_ID") & ". Assessment " & rsData.Fields("Assessment_ID")
    
        'no assessment id, do nothing
        Case Is = 0

        Case Is = 1
            'should never happen
            If IsNull(rsAsmnt.Fields("assessment_key")) Then
                Print #qc_log, "Null Assessment_ID in record #" & rsData.AbsolutePosition
            End If
            
            'check for mismatched WBID
            If Not (Trim(rsWb.Fields("waterbody_id")) = Trim(rsAsmnt.Fields("waterbody_id"))) Then
                Print #qc_log, "Mismatched assessment record error in waterbody_id.  Sampling_ID " & rsData.Fields("Sampling_ID")
            End If
 
            'check for mismatched source
            If Not (Trim(rsAsmnt.Fields("source")) = Trim(rsData.Fields("Source"))) Then
                Print #qc_log, "Mismatched assessment record error in source.  Sampling_ID " & rsData.Fields("Sampling_ID")
            End If
    
            'check for mismatched method
            If Not (rsMd.Fields("lookup_method_id") = (rsAsmnt.Fields("lookup_method_id"))) Then
                Print #qc_log, "Mismatched assessment record error in source.  Sampling_ID " & rsData.Fields("Sampling_ID")
            End If
            
            'check for mismatched start date
            If Not (rsAsmnt.Fields("start_date") = (rsData.Fields("start_date"))) Then
                Print #qc_log, "Mismatched assessment record error in start date.  Sampling_ID " & rsData.Fields("Sampling_ID")
            End If
            
            'check for mismatched end date
            If Not (rsAsmnt.Fields("end_date") = (rsData.Fields("end_date"))) Then
                Print #qc_log, "Mismatched assessment record error in end date.  Sampling_ID " & rsData.Fields("Sampling_ID")
            End If
            
            'open waterbody table to get related waterbody id
            rsTemp.CursorType = adOpenKeyset
            rsTemp.LockType = adLockOptimistic
            rsTemp.Open "SELECT MOF_waterbody_id FROM ffsbc.waterbody WHERE waterbody_id = " & rsAsmnt.Fields("waterbody_id") & ";", conn
            
            If rsTemp.RecordCount = 0 Then
                Print #qc_log, "Mismatched or missing waterbody name in assessment table.  Sampling_ID " & rsData.Fields("Sampling_ID")
            End If
            
            rsTemp.Close
            
            'check for mismatched project_ID
            'rsProj_Asmnt.CursorType = adOpenKeyset
            'rsProj_Asmnt.LockType = adLockOptimistic
            'rsProj_Asmnt.Open "SELECT * FROM ffsbc.project_assessment WHERE project_id = " & rsProj.Fields("project_id") & " AND assessment_id = " & a_id & ";", conn
            
            'If rsProj_Asmnt.RecordCount = 0 Then
              'Print #qc_log, "Mismatched project_id associated with assessment currently in database.  Sampling_ID " & rsData.Fields("Sampling_ID")
            'End If
            
            'rsProj_Asmnt.Close
            
    End Select
    
    'close handles to commit changes
    rsProj.Close
    rsMd.Close
    rsWb.Close
    rsAsmnt.Close
    rsSp.Close

    '*****************************finished processing current data record
    'set current record in Sampling Summary data to next record
    rsData.MoveNext
Loop

'close handles to commit changes
rsData.Close

'close logfiles
Close #wb_log
Close #qc_log
conn.Close

Exit_qc_sampling_summary:
    DoCmd.SetWarnings True
    Exit Sub

qc_sampling_summary_Err:
    MsgBox Err.Description
    Resume Exit_qc_sampling_summary
End Sub
