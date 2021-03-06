Attribute VB_Name = "qc_biological_data_8"
Private Sub qc_biological()
'qc entries for data and checks there is a match in lookup tables.  If there are multiple related waterbodies (i.e. creeks) this is logged
'assumes dates are valid, pH valid
'errors are logged to: qc_biological_log

On Error GoTo qc_biological_Err

Dim db As DAO.Database
Dim rsData As DAO.Recordset             'data to load

Dim rsWb As New ADODB.Recordset         'waterbody table
Dim rsAsmnt As New ADODB.Recordset      'assessment table
Dim rsTemp As New ADODB.Recordset       'miscellaneous recordset

Dim conn As New ADODB.Connection        'connection to small lakes
Dim cmd As New ADODB.Command

Dim wb_id As Variant                    'waterbody id
Dim r_id As Variant                     'region id
Dim a_id As Variant                     'assessment_id
Dim p_id As Variant                     'project_id

Dim fc_id As Variant                    'related fish_collection_id
Dim sp_id As Variant                    'species_id
Dim cl_id As Variant                    'clip_id
Dim ss_id As Variant                    'strain_species_id
Dim mt_id As Variant                    'lookup_maturity id
Dim m_id As Variant                     'lookup_method_id
Dim ms_id As Variant                    'lookup_mesh_size_id
Dim sc_id As Variant                    'lookup_stomach_content_id
Dim am_id As Variant                    'lookup_age_method_id

Dim a_key As Variant                    'assessment_key

Dim outfilepath As String               'filename for output

outfilepath = "\\FFSCOFP04\Users$\bonnie.robert\My Documents\Projects\DB Restructuring\Logfiles\qc_biological_log.txt"
qc_bio_log = FreeFile()
Close #qc_bio_log
Open outfilepath For Output As #qc_bio_log
Print #qc_bio_log, "Starting QC error report: " & Now

'open connection to small lakes test database'
conn.Open "SMALL_LAKES-TEST", "GOFISHBC/Bonnie.Robert"

'open Sampling Summary table in local Access instance
Set db = CurrentDb

'for table in access
'strSQL = "SELECT * FROM ffsbc_Biological_Data;"
strSQL = "select * from [Text;;FMT=Delimited;HDR=YES;IMEX=2;DATABASE=C:\Users\bonnie.robert\Desktop\New\S1501-2016_2017_Whiteswan_Winter_Creel].[Biological Data#csv]"

Set rsData = db.OpenRecordset(strSQL)
'for each record in Biological_Data dataset (read only)
Do Until rsData.EOF

    Debug.Print ("QC'ing Bio_Data_ID: " & rsData.Fields("Bio_Data_ID"))
    
    If (rsData.Fields("Date") > Date) Then
        Print #qc_bio_log, "Date error associated with biological data ID " & rsData.Fields("Bio_Data_ID")
    End If
    
    If (IsNull(rsData.Fields("Date"))) Then
            dat = "00000"
    Else:   dat = CStr(CLng(rsData.Fields("Date")))
    End If
                        
    'open method lookup table to get related lookup id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.lookup_method WHERE method_code = '" & Trim(rsData.Fields("Method")) & "';", conn
    If Not (IsNull(rsData.Fields("Method"))) Then
        If (rsTemp.RecordCount = 0) Then
            Print #qc_bio_log, "Method error associated with Bio_Data_ID " & rsData.Fields("Bio_Data_ID")
        Else:
            m_id = rsTemp.Fields("lookup_method_id")
        End If
        
    End If
    rsTemp.Close
                                    
    'open mesh size lookup table to get related lookup id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.lookup_mesh_size WHERE mesh_size = CAST('" & Trim(rsData.Fields("Net_Mesh_Size")) & "' AS Float);", conn
    
    If Not (IsNull(rsData.Fields("Net_Mesh_Size"))) Then
        If (rsTemp.RecordCount = 0) Then
            Print #qc_bio_log, "Mesh Size error associated with Bio_Data_ID " & rsData.Fields("Bio_Data_ID")
        Else:
            ms_id = rsTemp.Fields("lookup_mesh_size_id")
        End If
        
    End If
    rsTemp.Close
                                    
    'open stomach_contents lookup table to get related lookup id
    'rsTemp.CursorType = adOpenKeyset
    'rsTemp.LockType = adLockOptimistic
    'rsTemp.Open "SELECT * FROM ffsbc.lookup_stomach_content WHERE stomach_content_code = '" & Trim(rsData.Fields("Stomach_Contents")) & "';", conn
    
    'If Not (IsNull(rsData.Fields("Stomach_Contents"))) Then
        'If (rsTemp.RecordCount = 0) Then
            'Print #qc_bio_log, "stomach content error associated with Bio_Data_ID " & rsData.Fields("Bio_Data_ID")
        'Else:
            'sc_id = rsTemp.Fields("lookup_stomach_content_id")
        'End If
        
    'End If
    'rsTemp.Close
        
    'open age_method lookup table to get related lookup id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.lookup_age_method WHERE age_method = '" & Trim(rsData.Fields("Aged_Method")) & "';", conn
    
    If Not (IsNull(rsData.Fields("Aged_Method"))) Then
        If (rsTemp.RecordCount = 0) Then
            Print #qc_bio_log, "Aged Method error associated with Bio_Data_ID " & rsData.Fields("Bio_Data_ID")
        Else:
            am_id = rsTemp.Fields("lookup_age_method_id")
        End If
    
    End If
    rsTemp.Close
    
    'open maturity lookup table to get related lookup id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.lookup_maturity WHERE maturity_code = '" & Trim(rsData.Fields("Maturity")) & "';", conn
    
    If Not (IsNull(rsData.Fields("Maturity"))) Then
        If (rsTemp.RecordCount = 0) Then
            Print #qc_bio_log, "Maturity error associated with Bio_Data_ID " & rsData.Fields("Bio_Data_ID")
        Else:
            mt_id = rsTemp.Fields("lookup_maturity_id")
        End If

    End If
    rsTemp.Close
                                    
    'open strain_species lookup table to get related lookup id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.lookup_strain_species WHERE strain_species_code = '" & Trim(rsData.Fields("Strain")) & "';", conn
    
    If Not (IsNull(rsData.Fields("Strain"))) Then
        If (rsTemp.RecordCount = 0) Then
            Print #qc_bio_log, "Strain error associated with Bio_Data_ID " & rsData.Fields("Bio_Data_ID")
        Else:
            ss_id = rsTemp.Fields("lookup_strain_species_id")
        End If

    End If
    rsTemp.Close
    
    'open clip lookup table to get related lookup id
    'rsTemp.CursorType = adOpenKeyset
    'rsTemp.LockType = adLockOptimistic
    'rsTemp.Open "SELECT * FROM ffsbc.lookup_clip WHERE clip_code = '" & Trim(rsData.Fields("Clip")) & "';", conn

    'If Not (IsNull(rsData.Fields("Clip"))) Then
        'If (rsTemp.RecordCount = 0) Then
            'Print #qc_bio_log, "Clip error associated with Bio_Data_ID " & rsData.Fields("Bio_Data_ID")
        'Else:
            'cl_id = rsTemp.Fields("lookup_clip_id")
        'End If

    'End If
    'rsTemp.Close
                                                       
    'open species lookup table to get related lookup id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.species WHERE species_code = '" & Trim(rsData.Fields("Species")) & "';", conn
    
    If Not (IsNull(rsData.Fields("Species"))) Then
        If (rsTemp.RecordCount = 0) Then
            Print #qc_bio_log, "Species error associated with Bio_Data_ID " & rsData.Fields("Bio_Data_ID")
        Else:
            sp_id = rsTemp.Fields("species_id")
        End If

    End If
    rsTemp.Close

    'open waterbody table to get related waterbody id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.waterbody WHERE MOF_waterbody_id = '" & Trim(rsData.Fields("WBID")) & "';", conn
                      
    Select Case rsTemp.RecordCount
        'if no related waterbody, throw an error
        Case Is = 0
            Print #qc_bio_log, "No waterbody associated with Bio_Data_ID " & rsData.Fields("Bio_Data_ID")
        'if multiple related project, throw an error
        Case Is > 1
            Print #qc_bio_log, "Multiple waterbody associated with Bio_Data_ID " & rsData.Fields("Bio_Data_ID")
            
        'check that the waterbody name matches the waterbody name in the database
        Case Is = 1
            If Not (IsNull(rsData.Fields("Name"))) Then
                wb_name = rsData.Fields("Name")
                wb_name = Replace(wb_name, " Lake", "")
                wb_name = Replace(wb_name, " Creek", "")
        
                If (InStr(1, LCase(rsTemp.Fields("gazetted_name")), LCase(wb_name)) = 0) Then
                    If (InStr(1, LCase(rsTemp.Fields("alias")), LCase(wb_name)) = 0) Then
                        Print #qc_bio_log, "waterbody name error associated with Bio_Data_ID " & rsData.Fields("Bio_Data_ID") & ". Name is " & rsData.Fields("Name")
                    End If
                End If
            End If
            wb_id = rsTemp.Fields("waterbody_id")
    End Select
    rsTemp.Close
    
    'open region table to get related region id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.region WHERE region_number = '" & rsData.Fields("Region") & "';", conn
                                   
    If Not (IsNull(rsData.Fields("Region"))) Then
        
        If (rsTemp.RecordCount = 0) Then
            Print #qc_bio_log, "Region error associated with Bio_Data_ID " & rsData.Fields("Bio_Data_ID")
        End If

    End If
    rsTemp.Close
 
    'open assessment table to see if there's an associated assessment
    If IsNull(rsData.Fields("Assessment_Id")) Then
        a_key = rsData.Fields("WBID") & "_" & dat & "_" & m_name
    Else: a_key = rsData.Fields("Assessment_ID")
    End If
    
    'open assessment table to get assessment id
    rsAsmnt.CursorType = adOpenKeyset
    rsAsmnt.LockType = adLockOptimistic
    rsAsmnt.Open "SELECT * FROM ffsbc.assessment WHERE assessment_key = '" & a_key & "';", conn
  
    'assessment_id found in assessment table
    Select Case rsAsmnt.RecordCount
        
        'more than one, throw error
        Case Is > 1
            Print #qc_bio_log, "Multiple assessments error. Assessment " & rsData.Fields("Assessment_ID")
    
        'no assessment id, do nothing
        Case Is = 0

        Case Is = 1

            'check for mismatched WBID
            If Not (wb_id = rsAsmnt.Fields("waterbody_id")) Then
                Print #qc_bio_log, "Mismatched assessment record error in waterbody_id.  Bio_Data_ID " & rsData.Fields("Bio_Data_ID")
            End If
    
            'check for mismatched method
            If Not (m_id = (rsAsmnt.Fields("lookup_method_id"))) Then
                Print #qc_bio_log, "Mismatched assessment record error in source.  Bio_Data_ID " & rsData.Fields("Bio_Data_ID")
            End If
            
            'check for mismatched start date
            If Not (rsAsmnt.Fields("start_date") < (rsData.Fields("Date")) < rsAsmnt.Fields("end_date")) Then
                Print #qc_bio_log, "Mismatched assessment record error in dates. Bio_Data_ID  " & rsData.Fields("Bio_Data_ID")
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

Exit_qc_biological:
    DoCmd.SetWarnings True
    Exit Sub

qc_biological_Err:
    MsgBox Err.Description
    Resume Exit_qc_biological
End Sub



