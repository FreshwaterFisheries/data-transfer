Attribute VB_Name = "load_tds_ph_updates_4"
Private Sub load_tds_ph_updates()

'loads load_tds_ph_updates data into assessment, waterbody profile and measurement tables
'assumes qc on entries has been completed and there is a match in lookup tables.  If there are multiple related waterbodies (i.e. creeks) this is logged
'and the first lake is used.
'assumes dates are valid, values are valid for the following measurements:
'errors are logged to: lake_profile_ld_log.txt

On Error GoTo load_tds_ph_updates_Err

Dim db As DAO.Database
Dim rsData As DAO.Recordset             'data to load

Dim rsWb As New ADODB.Recordset         'waterbody table
Dim rsWP As New ADODB.Recordset         'waterbody profile table
Dim rsWPM As New ADODB.Recordset        'waterbody profile measurements table
Dim rsAsmnt As New ADODB.Recordset      'assessment table
Dim rsProj_Asmnt As New ADODB.Recordset 'project_assessment table
Dim rsTemp As New ADODB.Recordset       'miscellaneous recordset

Dim conn As New ADODB.Connection        'connection to small lakes
Dim cmd As New ADODB.Command

Dim wb_id As Variant                    'waterbody id
Dim wp_id As Variant                    'waterbody profile id
Dim a_id As Variant                     'assessment id
Dim m_id As Variant                     'lookup_method id
Dim mmnt_id As Variant                  'measurement id
Dim r_id As Variant                     'region id

Dim dat As Variant                      'date string for assessment key
Dim m_name As Variant                   'method code

Dim null_proj_id As Integer             'null project id in project table
Dim null_m_id As Integer                'null method id in project table

Dim fldlist() As Variant                'list of measurement fields

Dim outfilepath As String               'filename for output

null_m_id = 23
null_proj_id = 88



outfilepath = "\\FFSCOFP04\Users$\bonnie.robert\My Documents\Projects\DB Restructuring\Logfiles\tds_ph_update_ld_log.txt"
ld_tds_log = FreeFile()
Close #ld_tds_log
Open outfilepath For Output As #ld_tds_log
Print #ld_tds_log, "Starting ld error report: " & Now

'open connection to small lakes test database
conn.Open "SMALL_LAKES-TEST", "GOFISHBC/Bonnie.Robert"

'open Sampling Summary table in local Access instance
Set db = CurrentDb
Set rsData = db.OpenRecordset("SELECT * FROM ffsbc_TDS_pH_Updates;")

'open waterbody_profile table for adding new records
rsWP.CursorType = adOpenKeyset
rsWP.LockType = adLockOptimistic
rsWP.Open "SELECT * FROM ffsbc.waterbody_profile;", conn

'open waterbody_profile_measurement table to add new records
rsWPM.CursorType = adOpenKeyset
rsWPM.LockType = adLockOptimistic
rsWPM.Open "SELECT * FROM ffsbc.waterbody_profile_measurement;", conn

'open project_assessment table table to add new records
rsProj_Asmnt.CursorType = adOpenKeyset
rsProj_Asmnt.LockType = adLockOptimistic
rsProj_Asmnt.Open "SELECT * FROM ffsbc.project_assessment;", conn
    
'open assessment table to new add records
rsAsmnt.CursorType = adOpenKeyset
rsAsmnt.LockType = adLockOptimistic
rsAsmnt.Open "SELECT * FROM ffsbc.assessment;", conn

'for each record in TDS_pH_Updates dataset (read only)
Do Until rsData.EOF
    
    Debug.Print ("Loading TDS_pH_Update_ID: " & rsData.Fields("TDS_pH_Update_ID"))
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
        Print #ld_tds_log, "No waterbody associated with TDS_pH_Updates_ID " & rsData.Fields("TDS_pH_Updates_ID")
    Else:
        wb_id = rsTemp.Fields("waterbody_id")
    End If
    rsTemp.Close
    
    'open region table to get related region id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.waterbody_region WHERE waterbody_id = '" & wb_id & "';", conn
                      
    If (rsTemp.RecordCount = 0) Then
        Print #ld_tds_log, "No region associated with TDS_pH_Updates_ID " & rsData.Fields("TDS_pH_Updates_ID")
    Else:
        r_id = rsTemp.Fields("region_id")
    End If
    rsTemp.Close
    
    'get the excel-style serial date string
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
                !date_updated = Date
                !comments = "Data from old database - TDS/pH updates table. " & rsData.Fields("Comments")
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

    'create new record in waterbody_profile and enter data for current record
    With rsWP
        .AddNew
        !assessment_id = a_id
        !start_date = rsData.Fields("Date")
        !UTM_Zone = rsData.Fields("Zone")
        !UTM_Northing = rsData.Fields("Northing")
        !UTM_Easting = rsData.Fields("Easting")
        !date_added = Date
        .Update
        wp_id = rsWP.Fields("waterbody_profile_id")
    End With
    
    'lookup list - relates incoming data columns to measurement table
    fldlist = Array("TDS", "pH", "pH", "Bottom pH", "Temperature", "Conductivity", "Secchi_Depth", "High Conductance", "Low Conductance")
    
    'for each measurement found in the current record
    For i = 6 To 14
        
    'if there is data for the measurement(i.e. not 0 or NULL), create a record in waterbody_profile
    If Not IsNull(rsData.Fields(i)) Then
                    
        'open waterbody_profile_measurement table to obtain wbid
        rsTemp.CursorType = adOpenKeyset
        rsTemp.LockType = adLockOptimistic
        rsTemp.Open "SELECT * FROM ffsbc.lookup_measurement WHERE measurement_name = '" & fldlist(i - 6) & "';", conn
        mmnt_id = rsTemp.Fields("measurement_id")
        rsTemp.Close
                
        With rsWPM
            .AddNew
            !waterbody_profile_id = wp_id
            !depth = 0
            !measurement_id = mmnt_id
            !lo_bound = rsData.Fields(i)
            !hi_bound = rsData.Fields(i)
            .Update
        End With
                    
    End If
        
    Next i
    
rsData.MoveNext
Loop


'close handles to commit changes
rsData.Close
rsWPM.Close
rsWP.Close
rsAsmnt.Close
rsProj_Asmnt.Close
conn.Close

Exit_tds_ph_updates:
    DoCmd.SetWarnings True
    Exit Sub

load_tds_ph_updates_Err:
    MsgBox Err.Description
    Resume Exit_tds_ph_updates
End Sub
