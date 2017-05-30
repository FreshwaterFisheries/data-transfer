Attribute VB_Name = "load_sampling_summary_1"
Option Compare Database

Private Sub load_sampling_summary()
'loads sampling_summary data into waterbody profile and measurement tables
'assumes qc on entries has been completed and there is a match in project, and other lookup tables.  If there is a multiple project
'match, then 2 entries are added to project/assessment table.  If there are multiple related waterbodies (i.e. creeks) this is logged
'and the first creek is used.
'assumes dates are valid, pH valid
'errors are logged to:

On Error GoTo load_sampling_summary_Err

Dim db As DAO.Database
Dim rsData As DAO.Recordset             'data to load

Dim rsWb As New ADODB.Recordset         'waterbody table
Dim rsWP As New ADODB.Recordset         'waterbody profile table
Dim rsWPM As New ADODB.Recordset        'waterbody profile measurements table
Dim rsSp As New ADODB.Recordset         'species table
Dim rsReg As New ADODB.Recordset        'region table
Dim rsMmnt As New ADODB.Recordset       'measurement table
Dim rsAsmnt As New ADODB.Recordset      'assessment table
Dim rsProj As New ADODB.Recordset       'project table
Dim rsProj_Asmnt As New ADODB.Recordset 'project_assessment table
Dim rsHistCS As New ADODB.Recordset     'historical catch summary table
Dim rsMd As New ADODB.Recordset         'lookup method table
Dim conn As New ADODB.Connection        'connection to small lakes
Dim cmd As New ADODB.Command

Dim wb_id As Variant                    'waterbody id
Dim wp_id As Variant                    'waterbody profile id
Dim a_id As Variant                     'assessment id
Dim p_id As Variant                     'project id
Dim r_id As Variant                     'region id
Dim p_name As Variant                   'project_name
Dim m_id As Variant                     'method id
Dim mmnt_id As Variant                  'measurement id
Dim null_proj_id As Integer             'null project id in project table
Dim null_m_id As Integer               'null method id in project table
Dim null_sp_id As Integer               'null species id in project table
Dim fldlist() As Variant                'list of measurement fields
Dim d As Variant

Dim outfilepath As String               'filename for output

null_m_id = 23
null_proj_id = 88
null_sp_id = 62

fldlist = Array("TDS", "pH", "Temperature", "Conductivity")

outfilepath = "\\FFSCOFP04\Users$\bonnie.robert\My Documents\Projects\DB Restructuring\Logfiles\sampling_summary_load_log.txt"
ssld_log = FreeFile()
Close #ssld_log
Open outfilepath For Output As #ssld_log
Print #ssld_log, "Starting QC error report: " & Now

'open connection to small lakes test database
conn.Open "SMALL_LAKES-TEST", "GOFISHBC/Bonnie.Robert"

'open Sampling Summary table in local Access instance
Set db = CurrentDb
Set rsData = db.OpenRecordset("SELECT * FROM ffsbc_Sampling_Summary;")

'for each record in Sampling Summary dataset (read only)
Do Until rsData.EOF
    Debug.Print ("Loading Sampling_ID: " & rsData.Fields("Sampling_ID"))
    '******* get foreign key id's: project_id, method_id, species,_id, waterbody_id, region_id and (if exists) assessment_id

    'open project table to get related project id
    rsProj.CursorType = adOpenKeyset
    rsProj.LockType = adLockOptimistic
    rsProj.Open "SELECT project_id FROM ffsbc.project WHERE project_name = '" & Trim(rsData.Fields("FFSBC_ID")) & "';", conn
                
    If (rsProj.RecordCount = 0) Then
        p_id = null_proj_id
    Else:
        p_id = rsProj.Fields("project_id")
    End If
    rsProj.Close
                         
    'open method lookup table to get related lookup method id
    rsMd.CursorType = adOpenKeyset
    rsMd.LockType = adLockOptimistic
    rsMd.Open "SELECT lookup_method_id FROM ffsbc.lookup_method WHERE method_code = '" & Trim(rsData.Fields("Method")) & "';", conn
    
    If (rsMd.RecordCount = 0) Then
        m_id = null_m_id
    Else:
        m_id = rsMd.Fields("lookup_method_id")
    End If
    rsMd.Close
    
    'open species to get related species id
    rsSp.CursorType = adOpenKeyset
    rsSp.LockType = adLockOptimistic
    rsSp.Open "SELECT * FROM ffsbc.species WHERE species_code = '" & Trim(rsData.Fields("Species_Code")) & "';", conn
                      
    If (rsSp.RecordCount = 0) Then
        sp_id = null_sp_id
    Else:
        sp_id = rsSp.Fields("species_id")
    End If
    rsSp.Close
                                    
    'open waterbody table to get related waterbody id
    rsWb.CursorType = adOpenKeyset
    rsWb.LockType = adLockOptimistic
    rsWb.Open "SELECT * FROM ffsbc.waterbody WHERE MOF_waterbody_id = '" & Trim(rsData.Fields("WBID")) & "';", conn
                      
    If (rsWb.RecordCount = 0) Then
        Print #ssld, "No waterbody associated with assessment " & rsData.Fields("Assessment_ID")
    Else:
        wb_id = rsWb.Fields("waterbody_id")
    End If
    rsWb.Close
    
    'open region table to get related region id
    rsReg.CursorType = adOpenKeyset
    rsReg.LockType = adLockOptimistic
    rsReg.Open "SELECT * FROM ffsbc.waterbody_region WHERE waterbody_id = '" & wb_id & "';", conn
                      
    If (rsReg.RecordCount = 0) Then
        Print #ssld, "No region associated with assessment " & rsData.Fields("Assessment_ID")
    Else:
        r_id = rsReg.Fields("region_id")
    End If
    rsReg.Close
    
    'open assessment table to get assessment id
    rsAsmnt.CursorType = adOpenKeyset
    rsAsmnt.LockType = adLockOptimistic
    rsAsmnt.Open "SELECT * FROM ffsbc.assessment WHERE assessment_key = '" & Trim(rsData.Fields("Assessment_ID")) & "';", conn
    
    If (rsAsmnt.RecordCount = 0) Then
        a_id = 0
    Else:
        a_id = rsAsmnt.Fields("assessment_id")
    End If
    rsAsmnt.Close
    
    'open waterbody_profile table for adding new records
    rsWP.CursorType = adOpenKeyset
    rsWP.LockType = adLockOptimistic
    rsWP.Open "SELECT * FROM ffsbc.waterbody_profile;", conn

    'open waterbody_profile_measurement table to add new records
    rsWPM.CursorType = adOpenKeyset
    rsWPM.LockType = adLockOptimistic
    rsWPM.Open "SELECT * FROM ffsbc.waterbody_profile_measurement;", conn

    'open historical sampling summary table to add new records
    rsHistCS.CursorType = adOpenKeyset
    rsHistCS.LockType = adLockOptimistic
    rsHistCS.Open "SELECT * FROM ffsbc.historical_catch_summary;", conn

    'open project_assessment table table to add new records
    rsProj_Asmnt.CursorType = adOpenKeyset
    rsProj_Asmnt.LockType = adLockOptimistic
    rsProj_Asmnt.Open "SELECT * FROM ffsbc.project_assessment;", conn
    
    'open assessment table to add records assessment id
    rsAsmnt.CursorType = adOpenKeyset
    rsAsmnt.LockType = adLockOptimistic
    rsAsmnt.Open "SELECT * FROM ffsbc.assessment WHERE assessment_key = '" & Trim(rsData.Fields("Assessment_ID")) & "';", conn
    
    'add assessment and assessment/project record if none in assessment table
    If (a_id = 0) Then
        With rsAsmnt
            .AddNew
                !waterbody_id = wb_id
                !Assessment_Key = rsData.Fields("Assessment_ID")
                !region_id = r_id
                !Source = rsData.Fields("Source")
                !start_date = rsData.Fields("Start_Date")
                !end_date = rsData.Fields("End_Date")
                !lookup_method_id = m_id
                !date_added = rsData.Fields("Date_Added")
                !date_updated = Date
                !comments = rsData.Fields("Comments")
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
     rsAsmnt.Close
     rsProj_Asmnt.Close
     
    'add new record to historical_catch _summary table
    With rsHistCS
        .AddNew
        !assessment_id = a_id
        !species_id = sp_id
        !combined_cpue_for_each_species = rsData.Fields("Combined_CPUE_for_each_species")
        !min_length = rsData.Fields("Min_Length")
        !max_length = rsData.Fields("Max_Length")
        !average_length = rsData.Fields("Average_Length")
        !ci_on_average_length = rsData.Fields("CI_on_average_length")
        !average_fcf = rsData.Fields("Average_FCF")
        !average_wr = rsData.Fields("Average_Wr")
        .Update
    End With
    rsHistCS.Close

    'create new record in waterbody_profile and enter data for current record
    With rsWP
        .AddNew
        !assessment_id = a_id
        !waterbody_id = wb_id
        !measurement_date = rsData.Fields("Start_Date")
        !date_added = rsData.Fields("Date_Added")
        .Update
    End With
    'obtain waterbody profile id
    wp_id = rsWP.Fields("waterbody_profile_id")
    rsWP.Close
        
    'for each measurement found in the current record
    For Each fld In fldlist
        
    'if there is data for the measurement(i.e. not 0 or NULL), create a record in waterbody_profile
    If Not IsNull(rsData.Fields(fld)) Then
                    
        'open waterbody_profile_measurement table to obtain wbid
        rsMmnt.CursorType = adOpenKeyset
        rsMmnt.LockType = adLockOptimistic
        rsMmnt.Open "SELECT * FROM ffsbc.lookup_measurement WHERE measurement_name = '" & fld & "';", conn
        mmnt_id = rsMmnt.Fields("measurement_id")
        rsMmnt.Close
                
        With rsWPM
            .AddNew
            !waterbody_profile_id = wp_id
            !depth = 0
            !measurement_id = mmnt_id
            !lo_bound = rsData.Fields(fld)
            !hi_bound = rsData.Fields(fld)
            .Update
        End With
                    
    End If
        
    Next fld
    rsWPM.Close
    
    '*****************************finished processing current data record
    'set current record in Sampling Summary data to next record
    rsData.MoveNext
Loop

'close handles to commit changes
rsData.Close
conn.Close

Exit_load_sampling_summary:
    DoCmd.SetWarnings True
    Exit Sub

load_sampling_summary_Err:
    MsgBox Err.Description
    Resume Exit_load_sampling_summary
End Sub

