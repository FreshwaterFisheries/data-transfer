Attribute VB_Name = "qc_net_summary_7"
Option Compare Database
Private Sub qc_net_summary()
'loads sampling_summary data into waterbody profile and measurement tables
'includes a corresponding entry in assessment_id ***This part needs to be addressed!

On Error GoTo qc_net_summary_Err

Dim db As DAO.Database
Dim rsData As DAO.Recordset             'data to load
Dim rsAsmnt As New ADODB.Recordset      'assessment table
Dim rsProj As New ADODB.Recordset       'project table
Dim rsProj_Asmnt As New ADODB.Recordset 'project_assessment table
Dim rsTemp As New ADODB.Recordset        'miscellaneous table

Dim conn As New ADODB.Connection        'connection to small lakes
Dim cmd As New ADODB.Command

Dim a_id As Integer                     'assessment id
Dim p_id As Integer                     'project id
Dim wb_id As Variant                    'waterbody id
Dim m_id As Integer                     'method id
Dim r_id As Integer                     'region id
Dim fc_id As Integer                    'net_summary_survey id
Dim sp_id As Integer                    'species_id
Dim sd_id As Integer                    'sample design id
Dim s_id As Integer                     'setting id
Dim h_id As Integer                     'habitat_id

outfilepath = "\\FFSCOFP04\Users$\bonnie.robert\My Documents\Projects\DB Restructuring\Logfiles\net_summary_method_err_qc_log.txt"
qc_net_log = FreeFile()
Close #qc_net_log
Open outfilepath For Output As #qc_net_log

Print #qc_net_log, "Starting QC error report: " & Now

'open connection to small lakes test database
conn.Open "SMALL_LAKES-TEST", "GOFISHBC/Bonnie.Robert"

'open Sampling Summary table in local Access instance
Set db = CurrentDb
Set rsData = db.OpenRecordset("SELECT * FROM ffsbc_Net_Summary;")

'for each record in Sampling Summary dataset (read only)
Do Until rsData.EOF
    Debug.Print "Processing Record: " & rsData.Fields("Gillnet_Summary_ID")

    'open project table to get related project id
    rsProj.CursorType = adOpenKeyset
    rsProj.LockType = adLockOptimistic
    rsProj.Open "SELECT * FROM ffsbc.project WHERE project_name = '" & Trim(rsData.Fields("FFSBC_ID")) & "';", conn
           
    'check for valid project in project table
    Select Case rsProj.RecordCount
    
        'if no related project, throw an error
        Case Is = 0
            p_id = 88
            Print #qc_net_log, "no project associated with Net_Summary_ID: " & rsData.Fields("Net_Summary_ID")
              
        'if multiple related project, throw an error
        Case Is > 1
            p_id = 88
            Print #qc_net_log, "Multiple project associated with Net_Summary_ID: " & rsData.Fields("Net_Summary_ID")

        Case Is = 1
            p_id = rsProj.Fields("project_id")
        
    End Select
    rsProj.Close
        
    'open waterbody table to get related waterbody id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.waterbody WHERE MOF_waterbody_id = '" & Trim(rsData.Fields("WBID")) & "';", conn
                                                   
    Select Case rsTemp.RecordCount
        'if no related waterbody, throw an error
        Case Is = 0
            Print #qc_net_log, "No waterbody associated with Net_Summary_ID: " & rsData.Fields("Net_Summary_ID")
            'if multiple related project, throw an error
            wb_id = 0
        Case Is > 1
            Print #qc_net_log, "Multiple waterbodies associated with Net_Summary_ID: " & rsData.Fields("Net_Summary_ID")
            wb_id = 0
        Case Is = 1
            wb_id = rsTemp.Fields("waterbody_id")
    End Select
    rsTemp.Close
    
    'open method lookup table to get related lookup method id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.lookup_method WHERE method_code = '" & Trim(rsData.Fields("Method")) & "';", conn
                      
    Select Case rsTemp.RecordCount
        'if no related method, throw an error
        Case Is = 0
            Print #qc_net_log, "No method associated with Net_Summary_ID: " & rsData.Fields("Net_Summary_ID")
            m_id = 23
            'if multiple related method, throw an error
        Case Is > 1
            m_id = 23
            'Print #qc_net_log,  "Multiple methods associated with Net_Summary_ID: " & rsData.Fields("Net_Summary_ID")
        Case Is = 1
            m_id = rsTemp.Fields("lookup_method_id")
    End Select
    rsTemp.Close
    
    'open sample design lookup table to get related lookup method id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.lookup_sample_design WHERE sample_design_code = '" & Trim(rsData.Fields("Sample_Design")) & "';", conn
                      
    Select Case rsTemp.RecordCount
        'if no related method, throw an error
        Case Is = 0
            'Print #qc_net_log, "No sample design associated with assessment " & rsData.Fields("Net_Summary_ID")
            sd_id = 7
            'if multiple related sample_designs, throw an error
        Case Is > 1
            sd_id = 7
            'Print #qc_net_log,  "Multiple sample_designs associated with assessment " & rsData.Fields("Net_Summary_ID")
        Case Is = 1
            sd_id = rsTemp.Fields("lookup_sample_design_id")
    End Select
    rsTemp.Close
    
    'open setting lookup table to get related lookup method id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.lookup_setting WHERE setting_code = '" & Trim(rsData.Fields("Setting")) & "';", conn
                      
    Select Case rsTemp.RecordCount
        'if no related setting, throw an error
        Case Is = 0
            'Print #qc_net_log, "No setting associated with Net_Summary_ID: " & rsData.Fields("Net_Summary_ID")
            s_id = 5
            'if multiple related settings, throw an error
        Case Is > 1
            s_id = 0
            Print #qc_net_log, "Multiple settings associated with Net_Summary_ID: " & rsData.Fields("Net_Summary_ID")
        Case Is = 1
            s_id = rsTemp.Fields("lookup_setting_id")
    End Select
    rsTemp.Close
    
    'open habitat lookup table to get related lookup method id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.lookup_habitat_type WHERE habitat_type_code = '" & Trim(rsData.Fields("Habitat")) & "';", conn
                      
    Select Case rsTemp.RecordCount
        'if no related method, throw an error
        Case Is = 0
            'Print #qc_net_log, "No habitat_type associated with assessment " & rsData.Fields("Net_Summary_ID")
            h_id = 5
            'if multiple related habitats, throw an error
        Case Is > 1
            h_id = 5
            Print #qc_net_log, "Multiple habitats associated with assessment " & rsData.Fields("Net_Summary_ID")
        Case Is = 1
            h_id = rsTemp.Fields("lookup_habitat_type_id")
    End Select
    rsTemp.Close
    
    'open species table to get related lookup species id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.species WHERE species_code = '" & Trim(rsData.Fields("Species_Code")) & "';", conn
    
    Select Case rsTemp.RecordCount
        'if no related species, throw an error
        Case Is = 0
            Print #qc_net_log, "No species/invalid species associated with Net_Summary_ID: " & rsData.Fields("Net_Summary_ID")
            sp_id = 62
            'if multiple related species, throw an error
        Case Is = 1
            sp_id = rsTemp.Fields("species_id")
        Case Is > 1
            sp_id = 62
            'Print #qc_net_log,  "Multiple species associated with Net_Summary_ID: " & rsData.Fields("Net_Summary_ID")
    End Select
    rsTemp.Close
    
    'open assessment table to get assessment id
    rsAsmnt.CursorType = adOpenKeyset
    rsAsmnt.LockType = adLockOptimistic
    rsAsmnt.Open "SELECT * FROM ffsbc.assessment WHERE assessment_key = '" & Trim(rsData.Fields("Assessment_ID")) & "';", conn
    
    'assessment_id found in assessment table
    Select Case rsAsmnt.RecordCount
        
        'more than one, throw error
        Case Is > 1
            Print #qc_net_log, "multiple assessment_id's found for ffsbc.assessment: " & rsData.Fields("Assessment_ID")
    
        'no assessment id, do nothing
        Case Is = 0

        Case Is = 1
            'should never happen
            If IsNull(rsAsmnt.Fields("assessment_id")) Then
                Print #qc_net_log, "Null Assessment_ID in record #" & rsData.AbsolutePosition
            'otherwise, get the foriegn key id's for assessment, project, region, method, and waterbody that
            'are associated with the assessment table
            Else: a_id = (rsAsmnt.Fields("assessment_id"))
            End If
            
            'check for mismatched WBID
            If Not (wb_id = rsAsmnt.Fields("waterbody_id")) Then
                Print #qc_net_log, "error in record " & rsData.Fields(0) & _
                "     assessment_id:" & rsData.Fields("Assessment_id") & _
                "     uploading WBID:" & rsData.Fields("WBID") & _
                "     database waterbody_id:" & rsAsmnt.Fields("waterbody_id")
            End If
 
            'check for mismatched Source
            If Not (Trim(rsAsmnt.Fields("source")) = Trim(rsData.Fields("Source"))) Then
                'Print #qc_net_log, "error in record " & rsData.Fields(0) & _
                '"     assessment_key:" & rsData.Fields("Assessment_id") & _
                '"     uploading source:" & rsData.Fields("Source") & _
                '"     database source:" & rsAsmnt.Fields("source")
            End If
    
            'check for mismatched method
            If Not (m_id = rsAsmnt.Fields("lookup_method_id")) Then
                Print #qc_net_log, "error in record " & rsData.Fields(0) & _
                "     assessment_key:" & rsData.Fields("Assessment_id") & _
                "     uploading Method:" & rsData.Fields("Method") & _
                "     database lookup_method_id:" & rsAsmnt.Fields("lookup_method_id")
            End If
        
            'check for mismatched project_ID
            rsProj_Asmnt.CursorType = adOpenKeyset
            rsProj_Asmnt.LockType = adLockOptimistic
            rsProj_Asmnt.Open "SELECT * FROM ffsbc.project_assessment WHERE project_id = " & p_id & " AND assessment_id = " & a_id & ";", conn
            
            If (rsProj_Asmnt.RecordCount = 0 And Not (IsNull(rsData.Fields("FFSBC_ID")))) Then
             
              'Print #qc_net_log, "error in record " & rsData.Fields(0) & _
              '"     assessment_key:" & rsData.Fields("Assessment_id") & _
              '"     is not related to project:" & rsData.Fields("FFSBC_ID") & _
              '"     in the database"
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

conn.Close

Exit_qc_net_summary:
    DoCmd.SetWarnings True
    Exit Sub

qc_net_summary_Err:
    MsgBox Err.Description
    Resume Exit_qc_net_summary
End Sub


