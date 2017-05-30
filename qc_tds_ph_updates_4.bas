Attribute VB_Name = "qc_tds_ph_updates_4"
Private Sub qc_tds_ph_updates()
'qc entries for data and checks there is a match in lookup tables.  If there are multiple related waterbodies (i.e. creeks) this is logged
'assumes dates are valid, pH valid
'errors are logged to: qc_lkpfl_log

On Error GoTo qc_tds_ph_updates_Err

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

Dim dat As Variant                      'date string for assessment key
Dim m_name As String                    'method code name
Dim fldlist() As Variant                'list of measurement fields

Dim outfilepath As String               'filename for output

null_m_id = 23
null_proj_id = 88

outfilepath = "\\FFSCOFP04\Users$\bonnie.robert\My Documents\Projects\DB Restructuring\Logfiles\tds_qc_log.txt"
qc_tds_log = FreeFile()
Close #qc_tds_log
Open outfilepath For Output As #qc_tds_log
Print #qc_tds_log, "Starting QC error report: " & Now

'open connection to small lakes test database
conn.Open "SMALL_LAKES-TEST", "GOFISHBC/Bonnie.Robert"

'open Sampling Summary table in local Access instance
Set db = CurrentDb
Set rsData = db.OpenRecordset("SELECT * FROM ffsbc_TDS_pH_Updates;")


'for each record in TDS_pH_Updates dataset (read only)
Do Until rsData.EOF
    Debug.Print ("QC'ing TDS_pH_Updates: " & rsData.Fields("TDS_PH_Update_ID"))
    
    'check for valid date
    If (rsData.Fields("Date") > Date) Then
        Print #qc_tds_log, "Date error associated with record " & rsData.Fields("TDS_PH_Update_ID")
    End If
    
    If (IsNull(rsData.Fields("Date"))) Then
            dat = "00000"
    Else:   dat = CStr(CLng(rsData.Fields("Date")))
    End If
                    
    'open method lookup table to get related lookup method id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.lookup_method WHERE method_code = '" & Trim(rsData.Fields("Method")) & "';", conn
    
    If (rsTemp.RecordCount = 0) Then
        m_id = null_m_id
        m_name = "UNK"
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
        Print #qc_tds_log, "No waterbody associated with record " & rsData.Fields("TDS_PH_Update_ID")
    Else:
        wb_id = rsTemp.Fields("waterbody_id")
    End If
    rsTemp.Close
    
    'open region table to check the region is correct
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.waterbody_region WHERE waterbody_id = '" & wb_id & "';", conn
                      
    If (rsTemp.RecordCount = 0) Then
        Print #qc_tds_log, "No region associated with waterbody " & rsData.Fields("WBID")
    End If
    rsTemp.Close
    
    'create assessment_key and check if one is already in the database
    'open assessment table to see if there's an associated assessment
    a_key = rsData.Fields("WBID") & "_" & dat & "_" & m_name
    
    'open assessment table to get assessment id
    rsAsmnt.CursorType = adOpenKeyset
    rsAsmnt.LockType = adLockOptimistic
    rsAsmnt.Open "SELECT * FROM ffsbc.assessment WHERE assessment_key = '" & a_key & "';", conn
  
    'assessment_id found in assessment table
    Select Case rsAsmnt.RecordCount
        
        'more than one, throw error
        Case Is > 1
            Print #qc_tds_log, "Multiple assessments error.  TDS_PH_Update_ID " & rsData.Fields("TDS_PH_Update_ID") & ". Assessment " & a_key
    
        'no assessment id, do nothing
        Case Is = 0

        Case Is = 1

            'check for mismatched WBID
            If Not (wb_id = rsAsmnt.Fields("waterbody_id")) Then
                Print #qc_tds_log, "Mismatched assessment record error in waterbody_id.  TDS_PH_Update_ID " & rsData.Fields("TDS_PH_Update_ID")
            End If
    
            'check for mismatched method
            If Not (m_id = (rsAsmnt.Fields("lookup_method_id"))) Then
                Print #qc_tds_log, "Mismatched assessment record error in method.  TDS_PH_Update_ID " & rsData.Fields("TDS_PH_Update_ID")
            End If
            
            'check for mismatched start date
            If Not (rsAsmnt.Fields("start_date") < (rsData.Fields("Date")) < rsAsmnt.Fields("end_date")) Then
                Print #qc_tds_log, "Mismatched assessment record error in dates. TDS_PH_Update_ID  " & rsData.Fields("TDS_PH_Update_ID")
            End If
            
            If Not (rsAsmnt.Fields("source") = (rsData.Fields("Source"))) Then
                Print #qc_tds_log, "Mismatched assessment record error in source.  TDS_PH_Update_ID: " & rsData.Fields("TDS_PH_Update_ID") & _
                      "  assessment_id: " & rsAsmnt.Fields("assessment_id")
            End If
         
    End Select
    rsAsmnt.Close
    
    '*****************************finished processing current data record
    'set current record in lake_profile data to next record
    rsData.MoveNext
Loop

'close handles to commit changes
rsData.Close
Close #qc_tds_log
conn.Close

Exit_qc_tds_ph_updates:
    DoCmd.SetWarnings True
    Exit Sub

qc_tds_ph_updates_Err:
    MsgBox Err.Description
    Resume Exit_qc_tds_ph_updates
End Sub



