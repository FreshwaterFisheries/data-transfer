Attribute VB_Name = "load_lakes_0"
Option Compare Database

Private Sub load_lakes()
'loads lake data into waterbody, historical_FDIS, historical_BEC, waterbody_region, waterbody_access and waterbody_dimensions tables.  As well,
'species observations are recorded in fish collection under a dummy assessment, and area_2, perimeter_2 are added as another waterbody dimensions record.
'Dummy assessments for lakes transfer data is in the form of WBID_00000_UK (for unknown method).  This is noted in the 'comments' section of the
'are assessment table.
'Assumes there is a match in region, and other lookup tables.
'waterbody dimensions, waterbody access, fish collection for each lake where other
'errors are logged to: lkld_log

On Error GoTo load_lakes_Err

Dim db As DAO.Database
Dim rsData As DAO.Recordset             'data to load

'insert new records into these tables
Dim rsWb As New ADODB.Recordset         'waterbody table
Dim rsHistFDIS As New ADODB.Recordset   'historical FDIS table
Dim rsHistBEC As New ADODB.Recordset    'historical BEC table
Dim rsWbA As New ADODB.Recordset        'waterbody access table
Dim rsWbDim As New ADODB.Recordset      'waterbody dimensions table
Dim rsWbR As New ADODB.Recordset        'waterbody_region table
Dim rsFC As New ADODB.Recordset         'fish collection table
Dim rsFCC As New ADODB.Recordset        'fish collection count table
Dim rsTemp As New ADODB.Recordset       'temporary recordset

'foreign key tables, lookups
Dim rsSp As New ADODB.Recordset         'species table
Dim rsReg As New ADODB.Recordset        'region table
Dim rsAsmnt As New ADODB.Recordset      'assessment table
Dim rsProj_Asmnt As New ADODB.Recordset 'project_assessment table

'connection variables
Dim conn As New ADODB.Connection        'connection to small lakes
Dim cmd As New ADODB.Command

'foreign key id's
Dim wb_id As Variant                    'new waterbody id
Dim a_id As Variant                     'new assessment id
Dim r_id As Variant                     'existing region id
Dim sp_id As Integer                    'exisitng species id
Dim wbt_id As Integer                   'existing waterbody_type_id
Dim fc_id As Integer                    'new fish collection id

Dim null_proj_id As Integer             'null project id
Dim null_sp_id As Integer               'null species id
Dim null_md_id As Integer               'null species id
Dim null_r_id As Integer                'null region id

Dim spstring As String                  'string of species from lakes table
Dim splist() As String                  'list of species

Dim outfilepath As String               'filename for output

Dim wba_data As Boolean                 'indicator if data in waterbody_access
Dim wbd_data As Boolean                 'indicator if data in waterbody_access
Dim wbsp_data As Boolean                'indicator if data in species

null_proj_id = 88
null_sp_id = 62
null_md_id = 26
null_r_id = 13

outfilepath = "\\FFSCOFP04\Users$\bonnie.robert\My Documents\Projects\DB Restructuring\Logfiles\lakes_load_log.txt"
lkld_log = FreeFile()
Close #lkld_log
Open outfilepath For Output As #lkld_log
Print #lkld_log, "Starting QC error report: " & Now

  'open connection to small lakes test database
  conn.Open "SMALL_LAKES-TEST", "GOFISHBC/Bonnie.Robert"

  'open Lakes table in local Access instance
  Set db = CurrentDb
  Set rsData = db.OpenRecordset("SELECT * FROM ffsbc_Lakes;")

  '******* open tables for adding records
    'open waterbody table to retrieve/add new waterbody id
    rsWb.CursorType = adOpenKeyset
    rsWb.LockType = adLockOptimistic
    rsWb.Open "SELECT * FROM ffsbc.waterbody;", conn
    
    'open waterbody_access table for adding new records
    rsWbA.CursorType = adOpenKeyset
    rsWbA.LockType = adLockOptimistic
    rsWbA.Open "SELECT * FROM ffsbc.waterbody_access;", conn

    'open waterbody_dimensions table to add new records
    rsWbDim.CursorType = adOpenKeyset
    rsWbDim.LockType = adLockOptimistic
    rsWbDim.Open "SELECT * FROM ffsbc.waterbody_dimensions;", conn

    'open waterbody_region table to add new records
    rsWbR.CursorType = adOpenKeyset
    rsWbR.LockType = adLockOptimistic
    rsWbR.Open "SELECT * FROM ffsbc.waterbody_region;", conn

    'open historical FDIS table to add new records
    rsHistFDIS.CursorType = adOpenKeyset
    rsHistFDIS.LockType = adLockOptimistic
    rsHistFDIS.Open "SELECT * FROM ffsbc.historical_FDIS;", conn

    'open historical BEC table to add new records
    rsHistBEC.CursorType = adOpenKeyset
    rsHistBEC.LockType = adLockOptimistic
    rsHistBEC.Open "SELECT * FROM ffsbc.historical_BEC;", conn

    'open assessment table to add new assessment id
    rsAsmnt.CursorType = adOpenKeyset
    rsAsmnt.LockType = adLockOptimistic
    rsAsmnt.Open "SELECT * FROM ffsbc.assessment;", conn

    'open project_assessment table to add new records
    rsProj_Asmnt.CursorType = adOpenKeyset
    rsProj_Asmnt.LockType = adLockOptimistic
    rsProj_Asmnt.Open "SELECT * FROM ffsbc.project_assessment;", conn
    
    'open fish_collection table to add new records
    rsFC.CursorType = adOpenKeyset
    rsFC.LockType = adLockOptimistic
    rsFC.Open "SELECT * FROM ffsbc.fish_collection;", conn
    
    'open fish_collection_count table to add new records
    rsFCC.CursorType = adOpenKeyset
    rsFCC.LockType = adLockOptimistic
    rsFCC.Open "SELECT * FROM ffsbc.fish_collection_count;", conn

  'for each record in Lakes dataset (read only)
  Do Until rsData.EOF
    Debug.Print ("Loading SLD_ID: " & rsData.Fields("SLD_ID"))

    '**************************************get existing foreign key id's: region_id***********
    'open region table to get related region id
    rsTemp.CursorType = adOpenKeyset
    rsTemp.LockType = adLockOptimistic
    rsTemp.Open "SELECT * FROM ffsbc.region WHERE region_number = '" & rsData.Fields("Region") & "';", conn
                      
    If (rsTemp.RecordCount = 0) Then
        Print #lkld_log, "No region associated with WBID " & rsData.Fields("WBID")
        r_id = null_r_id
    Else:
        r_id = rsTemp.Fields("region_id")
    End If
    rsTemp.Close

    'set waterbody_type
    If (Left(rsData.Fields("WBID"), 5) = "00000") Then
        wbt_id = 2
    Else: wbt_id = 1
    End If

    '***************************************************add waterbody record*******************
    With rsWb
        .AddNew
        !waterbody_type_id = wbt_id
        !Watershed_Code = rsData.Fields("Watershed_Code")
        !MOF_waterbody_id = rsData.Fields("WBID")
        !Loc_FFSBC_Key = rsData.Fields("Loc_FFSBC_Key")
        !Loc_ID = rsData.Fields("Loc_ID")
        !Gazetted_Name = rsData.Fields("Gazetted_Name")
        !ALIAS = rsData.Fields("Alias")
        !Nearest_Town = rsData.Fields("Nearest_Town")
        !UTM_Zone = rsData.Fields("UTM_Zone")
        !UTM_Easting = rsData.Fields("UTM_Easting")
        !UTM_Northing = rsData.Fields("UTM_Northing")
        !date_added = rsData.Fields("Date_Added")
        .Update
        wb_id = rsWb.Fields("waterbody_id")
    End With
   
    '***********************************************add waterbody_region record*******************
    With rsWbR
        .AddNew
        !waterbody_id = wb_id
        !region_id = r_id
        .Update
    End With
    
    '***********************************************add historical_BEC record*******************
    'Add a new record if one of the BEC fields is not null
    If Not (IsNull(rsData.Fields("BEC_ZONE")) And _
            IsNull(rsData.Fields("BEC_SUB")) And _
            IsNull(rsData.Fields("BEC_VAR"))) Then
            
        With rsHistBEC
            .AddNew
            !waterbody_id = wb_id
            !BEC_ZONE = rsData.Fields("BEC_ZONE")
            !BEC_SUB = rsData.Fields("BEC_SUB")
            !BEC_VAR = rsData.Fields("BEC_VAR")
            .Update
        End With
    End If
  
    '***********************************************add historical_FDIS record*******************
    'Add a new record if one of the FDIS fields is not null
    If Not (IsNull(rsData.Fields("FDIS_ROAD_TO_LAKESHORE_DISTANCE")) And _
            IsNull(rsData.Fields("FDIS_OFF_ROAD_SITE_TO_LAKESHORE_DIS")) And _
            IsNull(rsData.Fields("FDIS_TRAIL_TO_LAKESHORE_DISTANCE")) And _
            IsNull(rsData.Fields("RD_CLOSEST_KM"))) Then
            
        With rsHistFDIS
            .AddNew
            !waterbody_id = wb_id
            !FDIS_ROAD_TO_LAKESHORE_DISTANCE = rsData.Fields("FDIS_ROAD_TO_LAKESHORE_DISTANCE")
            !FDIS_OFF_ROAD_SITE_TO_LAKESHORE_DIS = rsData.Fields("FDIS_OFF_ROAD_SITE_TO_LAKESHORE_DIS")
            !FDIS_TRAIL_TO_LAKESHORE_DISTANCE = rsData.Fields("FDIS_TRAIL_TO_LAKESHORE_DISTANCE")
            !RD_CLOSEST_KM = rsData.Fields("RD_CLOSEST_KM")
            .Update
        End With
    End If
    
     
   '********************* add assessment and assessment/project record if none in assessment table*******
   With rsAsmnt
        .AddNew
            !waterbody_id = wb_id
            !Assessment_Key = rsData.Fields("WBID") + "_00000_UP"
            !region_id = r_id
            !Source = "old SLD database. 1 - Lakes"
            !start_date = Date
            !end_date = Date
            !lookup_method_id = null_md_id
            !date_added = Date
            !date_updated = Date
            !comments = "data transferred from old database '1 - Lakes' table"
        .Update
        a_id = rsAsmnt.Fields("assessment_id")
    End With

    '********************************** add assessment_id and project_id to assessment_project table
    With rsProj_Asmnt
        .AddNew
            !project_id = null_proj_id
            !assessment_id = a_id
        .Update
    End With
  
  '*************************************** waterbody access table **********************
    With rsWbA
        .AddNew
            !assessment_id = a_id
            !DIRECTIONS = rsData.Fields("Directions")
            !Access_Info_Source = rsData.Fields("Access_Info_Source")
            !Access_Private = rsData.Fields("Access_Private")
            !Access = rsData.Fields("Access")
            !Boat_Trailer = rsData.Fields("Boat_Trailer")
            !Boat_Car_Topper = rsData.Fields("Boat_Car_Topper")
            !Source_Amenities = rsData.Fields("Source_Amenities")
            !Boat_Launches = rsData.Fields("Boat_Launches")
            !fishing_piers = rsData.Fields("Fishing_Piers")
            !Resorts = rsData.Fields("Resorts")
            !Campsites = rsData.Fields("Campsites")
            !DISTANCE = rsData.Fields("Distance")
            !date_added = rsData.Fields("Date_Added")
            !comments = rsData.Fields("Amenities_Comments")
        .Update
    End With

    '************************************* waterbody dimensions table ******************
    With rsWbDim
        .AddNew
            !assessment_id = a_id
            !Surface_Area_Comments = rsData.Fields("Surface_Area_Comments")
            !area_surface_ha = rsData.Fields("Area_Surface_ha")
            !depth_max_m = rsData.Fields("Depth_Max_m")
            !depth_mean_m = rsData.Fields("Depth_Mean_m")
            !elevation_m = rsData.Fields("Elevation_m")
            !perimeter_m = rsData.Fields("Perimeter_m")
            !area_littoral_ha = rsData.Fields("Area_Littoral_ha")
            !max_water_level = rsData.Fields("Max_Water_Level")
            !Area_Littoral_Percent = rsData.Fields("Area_Littoral_Percent")
            !no_of_outlets = rsData.Fields("No_of_Outlets")
            !no_of_inlets_permanent = rsData.Fields("No_of_Inlets_Permenant")
            !no_of_inlets_intermittent = rsData.Fields("No_of_Inlets_Intermittent")
            !resevoir_indicator = rsData.Fields("Reservoir_Indicator")
            !species_observed = rsData.Fields("Species_Observed")
            !species_observation_source_and_date = rsData.Fields("Species_Observation_Source_And_Date")
            !area_surface_2_ha = rsData.Fields("Area_Surface_2_ha")
            !perimeter_2_m = rsData.Fields("Perimeter_2_m")
            !source_area_perimeter = rsData.Fields("Source_Area_Perimeter")
            !source_lake_parameters = rsData.Fields("Source_Lake_Parameters")
            !date_lake_parameters = rsData.Fields("Date_Lake_Parameters")
            !date_added = rsData.Fields("Date_Added")
        .Update
    End With

    '******************************** open species table to find species
    If (Not IsNull(rsData.Fields("Species_Observed"))) Then
    
        splist = Split(rsData.Fields("Species_Observed"), ",")
        spstring = Join(splist, "','")
        
        rsSp.CursorType = adOpenKeyset
        rsSp.LockType = adLockOptimistic
        rsSp.Open "SELECT * FROM ffsbc.species WHERE species_code IN ('" & spstring & "');", conn
        
        'Add new record to fish_collection
        With rsFC
            .AddNew
            !assessment_id = a_id
            !lookup_method_id = null_md_id
            !date_added = rsData.Fields("Date_Added")
            !comments = "data transferred from old database '1 - Lakes' table"
            .Update
        End With
        fc_id = rsFC.Fields("fish_collection_id")

        'Loop thorugh species list - add species counts to fish_collection_counts table
        Do Until rsSp.EOF
            With rsFCC
                .AddNew
                !fish_collection_id = fc_id
                !species_id = rsSp.Fields("species_id")
                !Count = 1
                .Update
            End With
            rsSp.MoveNext
        Loop
        rsSp.Close
    End If
    
    '*****************************finished processing current data record
    'set current record in Sampling Summary data to next record
    rsData.MoveNext
Loop

'close handles to commit changes
rsWb.Close
rsWbR.Close
rsHistBEC.Close
rsHistFDIS.Close
rsAsmnt.Close
rsProj_Asmnt.Close
rsFC.Close
rsFCC.Close

rsData.Close
conn.Close

Exit_load_lakes:
    DoCmd.SetWarnings True
    Exit Sub

load_lakes_Err:
    MsgBox Err.Description
    Resume Exit_load_lakes
End Sub


