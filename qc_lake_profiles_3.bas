Attribute VB_Name = "qc_lake_profiles_3"
Private Sub qc_lake_profiles()
'qc entries for data and checks there is a match in lookup tables.  If there are multiple related waterbodies (i.e. creeks) this is logged
'assumes dates are valid, pH valid
'errors are logged to: qc_lkpfl_log

On Error GoTo qc_lake_profiles_Err

Dim db As DAO.Database
Dim rsData As DAO.Recordset             'data to load

Dim rsWb As New ADODB.Recordset         'waterbody table

Dim rsMmnt As New ADODB.Recordset       'measurement table
Dim rsAsmnt As New ADODB.Recordset      'assessment table
Dim rsProj_Asmnt As New ADODB.Recordset 'project_assessment table
Dim rsTemp As New ADODB.Recordset       'miscellaneous recordset

Dim conn As New ADODB.Connection        'connection to small lakes
Dim cmd As New ADODB.Command

Dim wb_id As Variant                    'waterbody id

Dim m_id As Variant                     'lookup_method id
Dim mmnt_id As Variant                  'measurement id
Dim r_id As Variant                     'region id

Dim null_proj_id As Integer             'null project id in project table
Dim null_m_id As Integer                'null method id in project table

Dim fldlist() As Variant                'list of measurement fields

Dim outfilepath As String               'filename for output

null_m_id = 23
null_proj_id = 88

outfilepath = "\\FFSCOFP04\Users$\bonnie.robert\My Documents\Projects\DB Restructuring\Logfiles\lake_profile_qc_log.txt"
qc_lkpfl_log = FreeFile()
Close #qc_lkpfl_log
Open outfilepath For Output As #qc_lkpfl_log
Print #qc_lkpfl_log, "Starting QC error report: " & Now

'open connection to small lakes test database
conn.Open "SMALL_LAKES-TEST", "GOFISHBC/Bonnie.Robert"

'open Sampling Summary table in local Access instance
Set db = CurrentDb
Set rsData = db.OpenRecordset("SELECT * FROM ffsbc_Lake_Profiles;")

'for each record in Lakes_Profiles dataset (read only)
Do Until rsData.EOF
    Debug.Print ("QC'ing Lake_Profile_ID: " & rsData.Fields("Lake_Profile_ID"))
    
    If (rsData.Fields("Date") > Date) Then
        Print #qc_lkpfl_log, "Date error associated with assessment " & rsData.Fields("Assessment_ID")
    End If
                         
    'open method lookup table to get related lookup method id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT lookup_method_id FROM ffsbc.lookup_method WHERE method_code = '" & Trim(rsData.Fields("Method")) & "';", conn
    
    If (rsTemp.RecordCount = 0) Then
        m_id = null_m_id
    Else:
        m_id = rsTemp.Fields("lookup_method_id")
    End If
    rsTemp.Close
                                    
    'open waterbody table to get related waterbody id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.waterbody WHERE MOF_waterbody_id = '" & Trim(rsData.Fields("WBID")) & "';", conn
                      
    If (rsTemp.RecordCount = 0) Then
        Print #qc_lkpfl_log, "No waterbody associated with Lake_Profile_ID: " & rsData.Fields("Lake_Profile_ID")
    Else:
        wb_id = rsTemp.Fields("waterbody_id")
    End If
    rsTemp.Close
    
    'open region table to get related region id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.waterbody_region WHERE waterbody_id = '" & wb_id & "';", conn
                      
    If (rsTemp.RecordCount = 0) Then
        Print #qc_lkpfl_log, "No region associated with Lake_Profile_ID " & rsData.Fields("Lake_Profile_ID")
    End If
    rsTemp.Close
    
    'open measurement table to obtain foriegn key measurement_id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
   
    If (rsData.Fields("Measure") = "pH" Or rsData.Fields("Measure") = "Turbidity") Then
            rsTemp.Open "SELECT * FROM ffsbc.lookup_measurement WHERE measurement_name = '" & rsData.Fields("Measure") & "';", conn
    Else:   rsTemp.Open "SELECT * FROM ffsbc.lookup_measurement WHERE measurement_name = '" & rsData.Fields("Measure") & _
                                                        "' AND measurement_units = '" & rsData.Fields("Units") & "';", conn
    End If
                                                        
    If (rsTemp.RecordCount = 0) Then
        Print #qc_lkpfl_log, "No measurement associated with Lake_Profile_ID: " & rsData.Fields("Lake_Profile_ID")
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
            Print #qc_lkpfl_log, "Multiple assessments error associated with Lake_Profile_ID: " & rsData.Fields("Lake_Profile_ID")
    
        'no assessment id, do nothing
        Case Is = 0

        Case Is = 1

            'check for mismatched WBID
            If Not (wb_id = rsAsmnt.Fields("waterbody_id")) Then
                Print #qc_lkpfl_log, "Mismatched assessment record error in waterbody_id.  Lake_Profile_ID " & rsData.Fields("Lake_Profile_ID")
            End If
    
            'check for mismatched method
            If Not (m_id = (rsAsmnt.Fields("lookup_method_id"))) Then
                Print #qc_lkpfl_log, "Mismatched assessment record error in source.  Lake_Profile_ID " & rsData.Fields("Lake_Profile_ID")
            End If
            
            'check for mismatched start date
            If Not (rsAsmnt.Fields("start_date") < (rsData.Fields("Date")) < rsAsmnt.Fields("end_date")) Then
                Print #qc_lkpfl_log, "Mismatched assessment record error in dates. Lake_Profile_ID  " & rsData.Fields("Lake_Profile_ID")
            End If
         
    End Select
    rsAsmnt.Close
    
    '*****************************finished processing current data record
    'set current record in lake_profile data to next record
    rsData.MoveNext
Loop

'close handles to commit changes
rsData.Close

conn.Close

Exit_qc_lake_profiles:
    DoCmd.SetWarnings True
    Exit Sub

qc_lake_profiles_Err:
    MsgBox Err.Description
    Resume Exit_qc_lake_profiles
End Sub

