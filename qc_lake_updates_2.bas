Attribute VB_Name = "qc_lake_updates_2"
Private Sub qc_lake_updates()
'qc entries for data and checks there is a match in lookup tables.  If there are multiple related waterbodies (i.e. creeks) this is logged
'assumes dates are valid, pH valid
'errors are logged to: qc_lkpfl_log

On Error GoTo qc_lake_updates_Err

Dim db As DAO.Database
Dim rsData As DAO.Recordset             'data to load

Dim rsWb As New ADODB.Recordset         'waterbody table
Dim rsAsmnt As New ADODB.Recordset      'assessment table
Dim rsTemp As New ADODB.Recordset       'miscellaneous recordset

Dim conn As New ADODB.Connection        'connection to small lakes
Dim cmd As New ADODB.Command

Dim wb_id As Variant                    'waterbody id
Dim m_id As Variant                     'lookup_method id
Dim r_id As Variant                     'region id
Dim a_id As Variant                     'assessment_id
Dim a_key As Variant                    'assessment_key
Dim dat As Variant                      'date string for assessment ID
Dim m_name As Variant                   'method name

Dim outfilepath As String               'filename for output

outfilepath = "\\FFSCOFP04\Users$\bonnie.robert\My Documents\Projects\DB Restructuring\Logfiles\lake_update_log.txt"
qc_lkup_log = FreeFile()
Close #qc_lkup_log
Open outfilepath For Output As #qc_lkup_log
Print #qc_lkup_log, "Starting QC error report: " & Now

'open connection to small lakes test database
conn.Open "SMALL_LAKES-TEST", "GOFISHBC/Bonnie.Robert"

'open Sampling Summary table in local Access instance
Set db = CurrentDb
Set rsData = db.OpenRecordset("SELECT * FROM ffsbc_Lake_Size_Updates;")

'for each record in Lakes_Profiles dataset (read only)
Do Until rsData.EOF

    Debug.Print ("QC'ing Lake_Size_Update_ID: " & rsData.Fields("Lake_Size_Update_ID"))
    
    If (rsData.Fields("Date") > Date) Then
        Print #qc_lkup_log, "Date error associated with lake_update.  Lake_Size_Update_ID:  " & rsData.Fields("Lake_Size_Update_ID")
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
        Print #qc_lkup_log, "Method error associated with lake_update. Lake_Size_Update_ID:  " & rsData.Fields("Lake_Size_Update_ID")
    Else:
        m_id = rsTemp.Fields("lookup_method_id")
        m_name = rsTemp.Fields("method_code")
    End If
    
    rsTemp.Close
                                    
    'open waterbody table to get related waterbody id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.waterbody WHERE MOF_waterbody_id = '" & Trim(rsData.Fields("WBID")) & "';", conn
    wb_id = rsTemp.Fields("waterbody_id")
                      
    Select Case rsTemp.RecordCount
        'if no related waterbody, throw an error
        Case Is = 0
            Print #qc_lkup_log, "No waterbody associated with lake_update.  Lake_Size_Update_ID:  " & rsData.Fields("Lake_Size_Update_ID")
        'if multiple related project, throw an error
        Case Is > 1
            Print #qc_lkup_log, "Multiple waterbody associated with lake_update.  Lake_Size_Update_ID:  " & rsData.Fields("Lake_Size_Update_ID")
            
        'check that the waterbody name matches the waterbody name in the database
        Case Is = 1
            If Not (IsNull(rsData.Fields("Name"))) Then
                wb_name = rsData.Fields("Name")
                wb_name = Replace(wb_name, " Lake", "")
                wb_name = Replace(wb_name, " Creek", "")
        
                If (InStr(1, LCase(rsTemp.Fields("gazetted_name")), LCase(wb_name)) = 0) Then
                    If (InStr(1, LCase(rsTemp.Fields("alias")), LCase(wb_name)) = 0) Then
                        Print #qc_lkup_log, "waterbody name error associated with lake_update " & rsData.Fields("Lake_Size_Update_ID")
                    End If
                End If
            End If

    End Select
    rsTemp.Close
    
    'open region table to get related region id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.region WHERE region_number = '" & rsData.Fields("Region") & "';", conn
                      
    If (rsTemp.RecordCount = 0) Then
        'Print #qc_lkup_log, "No region associated with lake_update " & rsData.Fields("Lake_Size_Update_ID")
    End If
    
    rsTemp.Close
    
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
            Print #qc_lkup_log, "Multiple assessments error. Lake_Size_Update_ID: " & rsData.Fields("Lake_Size_Update_ID")
    
        'no assessment id, do nothing
        Case Is = 0

        Case Is = 1

            'check for mismatched WBID
            If Not (wb_id = rsAsmnt.Fields("waterbody_id")) Then
                Print #qc_lkup_log, "Mismatched assessment record error in waterbody_id.  Lake_Size_Update_ID: " & rsData.Fields("Lake_Size_Update_ID")
            End If
    
            'check for mismatched method
            If Not (m_id = (rsAsmnt.Fields("lookup_method_id"))) Then
                Print #qc_lkup_log, "Mismatched assessment record error in source.  Lake_Size_Update_ID: " & rsData.Fields("Lake_Size_Update_ID")
            End If
            
            'check for mismatched start date
            If Not (rsAsmnt.Fields("start_date") < (rsData.Fields("Date")) < rsAsmnt.Fields("end_date")) Then
                Print #qc_lkup_log, "Mismatched assessment record error in dates. Lake_Size_Update_ID:  " & rsData.Fields("Lake_Size_Update_ID")
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

Exit_qc_lake_updates:
    DoCmd.SetWarnings True
    Exit Sub

qc_lake_updates_Err:
    MsgBox Err.Description
    Resume Exit_qc_lake_updates
End Sub
