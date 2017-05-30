Attribute VB_Name = "qc_creel_5"
Option Compare Database
Private Sub qc_creel()
'loads sampling_summary data into waterbody profile and measurement tables
'includes a corresponding entry in assessment_id ***This part needs to be addressed!

On Error GoTo qc_creel_Err

Dim db As DAO.Database
Dim rsData As DAO.Recordset             'data to load
Dim rsAsmnt As New ADODB.Recordset      'assessment table
Dim rsProj As New ADODB.Recordset       'project table
Dim rsProj_Asmnt As New ADODB.Recordset  'project_assessment table
Dim rsTemp As New ADODB.Recordset        'waterbody table

Dim conn As New ADODB.Connection        'connection to small lakes
Dim cmd As New ADODB.Command

Dim match As Boolean
Dim cntr As Integer

Dim a_id As Integer                     'assessment id
Dim p_id As Integer                     'project id
Dim wb_id As Variant                    'waterbody id
Dim m_id As Integer                     'method id
Dim r_id As Integer                     'region id
Dim cs_id As Integer                    'creel_survey id
Dim sp_id As Integer                    'species_id

outfilepath = "\\FFSCOFP04\Users$\bonnie.robert\My Documents\Projects\DB Restructuring\Logfiles\creel_qc_log.txt"
qc_crl_log = FreeFile()
Close #qc_crl_log
Open outfilepath For Output As #qc_crl_log
Print #qc_crl_log, "Starting QC error report: " & Now

'open connection to small lakes test database
conn.Open "SMALL_LAKES-TEST", "GOFISHBC/Bonnie.Robert"

'open Sampling Summary table in local Access instance
Set db = CurrentDb
Set rsData = db.OpenRecordset("SELECT * FROM ffsbc_Creel;")

'for each record in Creel dataset (read only)
Do Until rsData.EOF
    Debug.Print "Procesing Record: " & rsData.Fields("Creel_ID")

    'open project table to get related project id
    rsProj.CursorType = adOpenKeyset
    rsProj.LockType = adLockOptimistic
    rsProj.Open "SELECT * FROM ffsbc.project WHERE project_name = '" & Trim(rsData.Fields("FFSBC_ID")) & "';", conn
           
    'check for valid project in project table
    Select Case rsProj.RecordCount
    
        'if no related project, throw an error
        Case Is = 0
            p_id = 0
            Print #qc_crl_log, "no project associated with Creel_ID " & rsData.Fields("Creel_ID")
              
        'if multiple related project, throw an error
        Case Is > 1
            p_id = 0
            Print #qc_crl_log, "Multiple project associated with Creel_ID " & rsData.Fields("Creel_ID")

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
            Print #qc_crl_log, "No waterbody associated with Creel_ID " & rsData.Fields("Creel_ID")
            'if multiple related project, throw an error
            wb_id = 0
        Case Is > 1
            Print #qc_crl_log, "Multiple waterbodies associated with Creel_ID " & rsData.Fields("Creel_ID")
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
            Print #qc_crl_log, "No method associated with Creel_ID " & rsData.Fields("Creel_ID")
            m_id = 0
            'if multiple related method, throw an error
        Case Is > 1
            m_id = 0
            Print #qc_crl_log, "Multiple methods associated with Creel_ID " & rsData.Fields("Creel_ID")
        Case Is = 1
            m_id = rsTemp.Fields("lookup_method_id")
    End Select
    rsTemp.Close
    
    'open species table to get related lookup species id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.species WHERE species_code = '" & Trim(rsData.Fields("Target_Species")) & "';", conn
    
    Select Case rsTemp.RecordCount
        'if no related species, throw an error
        Case Is = 0
            Print #qc_crl_log, "No species/invalid species associated with Creel_ID " & rsData.Fields("Creel_ID")
            sp_id = 0
            'if multiple related species, throw an error
        Case Is = 1
            sp_id = rsTemp.Fields("species_id")
        Case Is > 1
            sp_id = 0
            Print #qc_crl_log, "Multiple species associated with Creel_ID " & rsData.Fields("Creel_ID")
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
            Print #qc_crl_log, "multiple assessment_id's found for Creel_ID: " & rsData.Fields("Creel_ID")
    
        'no assessment id, do nothing
        Case Is = 0

        Case Is = 1
            'should never happen
            If IsNull(rsAsmnt.Fields("assessment_id")) Then
                Print #qc_crl_log, "Null Assessment_ID in record #" & rsData.AbsolutePosition
            'otherwise, get the foriegn key id's for assessment, project, region, method, and waterbody that
            'are associated with the assessment table
            Else: a_id = (rsAsmnt.Fields("assessment_id"))
            End If
            
            'check for mismatched WBID
            If Not (wb_id = rsAsmnt.Fields("waterbody_id")) Then
                Print #qc_crl_log, "error in record " & rsData.Fields(0) & _
                "     assessment_id:" & rsData.Fields("Assessment_id") & _
                "     uploading WBID:" & rsData.Fields("WBID") & _
                "     database waterbody_id:" & rsAsmnt.Fields("waterbody_id")
            End If
 
            'check for mismatched Source
            If Not (Trim(rsAsmnt.Fields("source")) = Trim(rsData.Fields("Source"))) Then
                Print #qc_crl_log, "error in record " & rsData.Fields(0) & _
                "     assessment_key:" & rsData.Fields("Assessment_id") & _
                "     uploading source:" & rsData.Fields("Source") & _
                "     database source:" & rsAsmnt.Fields("source")
            End If
    
            'check for mismatched method
            If Not (m_id = rsAsmnt.Fields("lookup_method_id")) Then
                Print #qc_crl_log, "error in record " & rsData.Fields(0) & _
                "     assessment_key:" & rsData.Fields("Assessment_id") & _
                "     uploading Method:" & rsData.Fields("Method") & _
                "     database lookup_method_id:" & rsAsmnt.Fields("lookup_method_id")
            End If
        
            'check for mismatched project_ID
            rsProj_Asmnt.CursorType = adOpenKeyset
            rsProj_Asmnt.LockType = adLockOptimistic
            rsProj_Asmnt.Open "SELECT * FROM ffsbc.project_assessment WHERE project_id = " & p_id & " AND assessment_id = " & a_id & ";", conn
            
            If rsProj_Asmnt.RecordCount = 0 Then
              Print #qc_crl_log, "error in record " & rsData.Fields(0) & _
              "     assessment_key:" & rsData.Fields("Assessment_id") & _
              "     is not related to project:" & rsData.Fields("FFSBC_ID") & _
              "     in the database"
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

Exit_qc_creel:
    DoCmd.SetWarnings True
    Exit Sub

qc_creel_Err:
    MsgBox Err.Description
    Resume Exit_qc_creel
End Sub

