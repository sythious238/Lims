Attribute VB_Name = "modSeed"
Option Compare Database
Option Explicit

' =============================================================================
'  EnviroLIMS - reference and demo data
'
'  IMPORTANT: The MDL / RL / QC acceptance values below are EXAMPLES so the
'  system works out of the box. Replace them with your laboratory's own
'  method detection limit studies and QC criteria before reporting data.
'  Regulatory limits are US EPA National Primary/Secondary Drinking Water
'  values; confirm the limits that apply in your jurisdiction and permits.
'  Numbers are passed as strings so SQL is locale-independent.
' =============================================================================

Private Sub X(ByVal sql As String)
    CurrentDb.Execute sql, dbFailOnError
End Sub

' Executes an INSERT...SELECT and fails loudly if a lookup name was mistyped.
Private Sub XOne(ByVal sql As String, ByVal what As String)
    Dim db As DAO.Database
    Set db = CurrentDb
    db.Execute sql, dbFailOnError
    If db.RecordsAffected = 0 Then Err.Raise vbObjectError + 513, "modSeed", "Reference data lookup failed: " & what
End Sub

Private Function Q(ByVal s As String) As String
    Q = SqlText(s)
End Function

' ----------------------------------------------------------------------------
'  Reference data
' ----------------------------------------------------------------------------

Private Sub AddSetting(ByVal n As String, ByVal v As String, ByVal d As String)
    X "INSERT INTO tblSettings (SettingName, SettingValue, Description) VALUES (" & Q(n) & ", " & Q(v) & ", " & Q(d) & ")"
End Sub

Private Sub AddMatrix(ByVal code As String, ByVal matrixName As String, ByVal matrixClass As String)
    X "INSERT INTO tblMatrices (MatrixCode, MatrixName, MatrixClass) VALUES (" & Q(code) & ", " & Q(matrixName) & ", " & Q(matrixClass) & ")"
End Sub

Private Sub AddAnalyte(ByVal analyteName As String, ByVal cas As String, ByVal grp As String)
    X "INSERT INTO tblAnalytes (AnalyteName, CASNumber, AnalyteGroup) VALUES (" & Q(analyteName) & ", " & _
      IIf(Len(cas) = 0, "Null", Q(cas)) & ", " & Q(grp) & ")"
End Sub

Private Sub AddMethod(ByVal code As String, ByVal methodName As String, ByVal matrixClass As String, ByVal holdHours As String, _
                      ByVal preservation As String, ByVal container As String, ByVal tatDays As String)
    X "INSERT INTO tblMethods (MethodCode, MethodName, MatrixClass, HoldTimeHours, Preservation, ContainerType, TurnaroundDays, Active) " & _
      "VALUES (" & Q(code) & ", " & Q(methodName) & ", " & Q(matrixClass) & ", " & holdHours & ", " & Q(preservation) & ", " & _
      Q(container) & ", " & tatDays & ", True)"
End Sub

Private Sub AddMA(ByVal methodCode As String, ByVal analyteName As String, ByVal units As String, ByVal mdl As String, ByVal rl As String, _
                  ByVal recLow As String, ByVal recHigh As String, ByVal rpd As String)
    XOne "INSERT INTO tblMethodAnalytes (MethodID, AnalyteID, Units, MDL, RL, RecoveryLow, RecoveryHigh, RPDLimit) " & _
         "SELECT m.MethodID, a.AnalyteID, " & Q(units) & ", " & mdl & ", " & rl & ", " & recLow & ", " & recHigh & ", " & rpd & " " & _
         "FROM tblMethods AS m, tblAnalytes AS a WHERE m.MethodCode=" & Q(methodCode) & " AND a.AnalyteName=" & Q(analyteName), _
         methodCode & " / " & analyteName
End Sub

Private Sub AddLimit(ByVal limitSet As String, ByVal analyteName As String, ByVal matrixCode As String, ByVal limitValue As String, _
                     ByVal units As String, ByVal limitType As String, ByVal source As String)
    XOne "INSERT INTO tblRegulatoryLimits (LimitSet, AnalyteID, MatrixID, LimitValue, Units, LimitType, Source) " & _
         "SELECT " & Q(limitSet) & ", a.AnalyteID, mx.MatrixID, " & limitValue & ", " & Q(units) & ", " & Q(limitType) & ", " & Q(source) & " " & _
         "FROM tblAnalytes AS a, tblMatrices AS mx WHERE a.AnalyteName=" & Q(analyteName) & " AND mx.MatrixCode=" & Q(matrixCode), _
         limitSet & " / " & analyteName & " / " & matrixCode
End Sub

' Primary MCLs applied to both drinking water and groundwater.
Private Sub AddMCL(ByVal analyteName As String, ByVal limitValue As String, ByVal limitType As String)
    AddLimit "EPA Primary MCL", analyteName, "DW", limitValue, "mg/L", limitType, "40 CFR 141"
    AddLimit "EPA Primary MCL", analyteName, "GW", limitValue, "mg/L", limitType, "40 CFR 141 (applied to groundwater)"
End Sub

Public Sub SeedReferenceData()
    ' --- Settings ---
    AddSetting "LabName", "Your Environmental Laboratory", "Printed on report headers"
    AddSetting "LabAddress", "123 Laboratory Way, City, ST 00000", "Printed on report headers"
    AddSetting "LabPhone", "(555) 555-0100", "Printed on report headers"
    AddSetting "Accreditation", "Accreditation / certification number", "Printed on the Certificate of Analysis"
    AddSetting "WorkOrderPrefix", "WO", "Work order numbers are <prefix><yymm>-<seq>"
    AddSetting "LabDirector", "Laboratory Director", "Signatory on the Certificate of Analysis"

    ' --- Matrices ---
    AddMatrix "DW", "Drinking Water", "Aqueous"
    AddMatrix "GW", "Groundwater", "Aqueous"
    AddMatrix "SW", "Surface Water", "Aqueous"
    AddMatrix "WW", "Wastewater", "Aqueous"
    AddMatrix "SO", "Soil", "Solid"
    AddMatrix "SD", "Sediment", "Solid"
    AddMatrix "SL", "Sludge", "Solid"

    ' --- Analytes ---
    AddAnalyte "Arsenic", "7440-38-2", "Metals"
    AddAnalyte "Cadmium", "7440-43-9", "Metals"
    AddAnalyte "Chromium", "7440-47-3", "Metals"
    AddAnalyte "Copper", "7440-50-8", "Metals"
    AddAnalyte "Lead", "7439-92-1", "Metals"
    AddAnalyte "Selenium", "7782-49-2", "Metals"
    AddAnalyte "Mercury", "7439-97-6", "Metals"
    AddAnalyte "Nitrate as N", "14797-55-8", "Anions"
    AddAnalyte "Chloride", "16887-00-6", "Anions"
    AddAnalyte "Sulfate", "14808-79-8", "Anions"
    AddAnalyte "Fluoride", "16984-48-8", "Anions"
    AddAnalyte "pH", "", "General Chemistry"
    AddAnalyte "Total Suspended Solids", "", "General Chemistry"
    AddAnalyte "Biochemical Oxygen Demand (5-day)", "", "General Chemistry"
    AddAnalyte "Total Coliform", "", "Microbiology"
    AddAnalyte "E. coli", "", "Microbiology"
    AddAnalyte "Benzene", "71-43-2", "Volatile Organics"
    AddAnalyte "Toluene", "108-88-3", "Volatile Organics"
    AddAnalyte "Ethylbenzene", "100-41-4", "Volatile Organics"
    AddAnalyte "Xylenes (total)", "1330-20-7", "Volatile Organics"
    AddAnalyte "Trichloroethene", "79-01-6", "Volatile Organics"
    AddAnalyte "Tetrachloroethene", "127-18-4", "Volatile Organics"
    AddAnalyte "Vinyl chloride", "75-01-4", "Volatile Organics"

    ' --- Methods (hold time in hours, turnaround in days) ---
    AddMethod "EPA 200.8", "Metals by ICP-MS", "Aqueous", "4320", "HNO3 to pH<2", "250 mL HDPE", "10"
    AddMethod "EPA 245.1", "Mercury by CVAA", "Aqueous", "672", "HNO3 to pH<2", "250 mL HDPE", "10"
    AddMethod "EPA 300.0", "Inorganic Anions by Ion Chromatography", "Aqueous", "48", "Cool to <=6 C", "250 mL HDPE", "5"
    AddMethod "SM 4500-H+ B", "pH (Electrometric)", "Aqueous", "0.25", "None - analyze immediately", "250 mL HDPE", "1"
    AddMethod "SM 2540 D", "Total Suspended Solids", "Aqueous", "168", "Cool to <=6 C", "1 L HDPE", "7"
    AddMethod "SM 5210 B", "Biochemical Oxygen Demand (5-day)", "Aqueous", "48", "Cool to <=6 C", "1 L HDPE", "7"
    AddMethod "SM 9223 B", "Total Coliform / E. coli (Colilert)", "Aqueous", "30", "Na2S2O3, cool to <10 C", "120 mL sterile", "2"
    AddMethod "EPA 524.2", "Volatile Organics in Drinking Water (GC/MS)", "Aqueous", "336", "HCl to pH<2, cool", "3 x 40 mL VOA", "10"
    AddMethod "EPA 6010D", "Metals in Solids by ICP-OES", "Solid", "4320", "Cool to <=6 C", "4 oz glass jar", "10"
    AddMethod "EPA 8260D", "Volatile Organics in Soil (GC/MS)", "Solid", "336", "Methanol / cool to <=6 C", "Terracore kit", "10"

    ' --- Method analytes: units, MDL, RL, %Rec low, %Rec high, RPD (EXAMPLE VALUES) ---
    AddMA "EPA 200.8", "Arsenic", "mg/L", "0.0004", "0.001", "85", "115", "20"
    AddMA "EPA 200.8", "Cadmium", "mg/L", "0.0001", "0.0005", "85", "115", "20"
    AddMA "EPA 200.8", "Chromium", "mg/L", "0.0005", "0.002", "85", "115", "20"
    AddMA "EPA 200.8", "Copper", "mg/L", "0.0005", "0.002", "85", "115", "20"
    AddMA "EPA 200.8", "Lead", "mg/L", "0.0002", "0.001", "85", "115", "20"
    AddMA "EPA 200.8", "Selenium", "mg/L", "0.0005", "0.002", "85", "115", "20"
    AddMA "EPA 245.1", "Mercury", "mg/L", "0.00005", "0.0002", "85", "115", "20"
    AddMA "EPA 300.0", "Nitrate as N", "mg/L", "0.01", "0.05", "90", "110", "15"
    AddMA "EPA 300.0", "Chloride", "mg/L", "0.1", "0.5", "90", "110", "15"
    AddMA "EPA 300.0", "Sulfate", "mg/L", "0.2", "1", "90", "110", "15"
    AddMA "EPA 300.0", "Fluoride", "mg/L", "0.02", "0.1", "90", "110", "15"
    AddMA "SM 4500-H+ B", "pH", "SU", "Null", "Null", "Null", "Null", "Null"
    AddMA "SM 2540 D", "Total Suspended Solids", "mg/L", "1", "2", "80", "120", "20"
    AddMA "SM 5210 B", "Biochemical Oxygen Demand (5-day)", "mg/L", "2", "2", "85", "115", "20"
    AddMA "SM 9223 B", "Total Coliform", "MPN/100mL", "1", "1", "Null", "Null", "Null"
    AddMA "SM 9223 B", "E. coli", "MPN/100mL", "1", "1", "Null", "Null", "Null"
    AddMA "EPA 524.2", "Benzene", "mg/L", "0.0001", "0.0005", "70", "130", "20"
    AddMA "EPA 524.2", "Toluene", "mg/L", "0.0001", "0.0005", "70", "130", "20"
    AddMA "EPA 524.2", "Ethylbenzene", "mg/L", "0.0001", "0.0005", "70", "130", "20"
    AddMA "EPA 524.2", "Xylenes (total)", "mg/L", "0.0002", "0.0005", "70", "130", "20"
    AddMA "EPA 524.2", "Trichloroethene", "mg/L", "0.0001", "0.0005", "70", "130", "20"
    AddMA "EPA 524.2", "Tetrachloroethene", "mg/L", "0.0001", "0.0005", "70", "130", "20"
    AddMA "EPA 524.2", "Vinyl chloride", "mg/L", "0.0001", "0.0005", "70", "130", "20"
    AddMA "EPA 6010D", "Arsenic", "mg/kg", "0.2", "1", "80", "120", "20"
    AddMA "EPA 6010D", "Cadmium", "mg/kg", "0.05", "0.5", "80", "120", "20"
    AddMA "EPA 6010D", "Chromium", "mg/kg", "0.1", "1", "80", "120", "20"
    AddMA "EPA 6010D", "Copper", "mg/kg", "0.2", "1", "80", "120", "20"
    AddMA "EPA 6010D", "Lead", "mg/kg", "0.2", "1", "80", "120", "20"
    AddMA "EPA 8260D", "Benzene", "mg/kg", "0.001", "0.005", "70", "130", "30"
    AddMA "EPA 8260D", "Toluene", "mg/kg", "0.001", "0.005", "70", "130", "30"
    AddMA "EPA 8260D", "Ethylbenzene", "mg/kg", "0.001", "0.005", "70", "130", "30"
    AddMA "EPA 8260D", "Xylenes (total)", "mg/kg", "0.002", "0.005", "70", "130", "30"
    AddMA "EPA 8260D", "Trichloroethene", "mg/kg", "0.001", "0.005", "70", "130", "30"
    AddMA "EPA 8260D", "Tetrachloroethene", "mg/kg", "0.001", "0.005", "70", "130", "30"

    ' --- US EPA drinking water limits (mg/L) ---
    AddMCL "Arsenic", "0.010", "MCL"
    AddMCL "Cadmium", "0.005", "MCL"
    AddMCL "Chromium", "0.1", "MCL"
    AddMCL "Copper", "1.3", "Action Level"
    AddMCL "Lead", "0.015", "Action Level"
    AddMCL "Selenium", "0.05", "MCL"
    AddMCL "Mercury", "0.002", "MCL"
    AddMCL "Nitrate as N", "10", "MCL"
    AddMCL "Fluoride", "4.0", "MCL"
    AddMCL "Benzene", "0.005", "MCL"
    AddMCL "Toluene", "1", "MCL"
    AddMCL "Ethylbenzene", "0.7", "MCL"
    AddMCL "Xylenes (total)", "10", "MCL"
    AddMCL "Trichloroethene", "0.005", "MCL"
    AddMCL "Tetrachloroethene", "0.005", "MCL"
    AddMCL "Vinyl chloride", "0.002", "MCL"
    AddLimit "EPA Secondary SMCL", "Chloride", "DW", "250", "mg/L", "SMCL", "40 CFR 143"
    AddLimit "EPA Secondary SMCL", "Sulfate", "DW", "250", "mg/L", "SMCL", "40 CFR 143"
    AddLimit "EPA Secondary SMCL", "Fluoride", "DW", "2.0", "mg/L", "SMCL", "40 CFR 143"
    AddLimit "EPA Secondary SMCL", "Copper", "DW", "1.0", "mg/L", "SMCL", "40 CFR 143"
End Sub

' ----------------------------------------------------------------------------
'  Demo data (a fictional client with one work order part-way through analysis)
' ----------------------------------------------------------------------------

Private Sub AddDemoSample(ByVal woID As Long, ByVal projectID As Long, ByVal clientSampleID As String, ByVal locationCode As String, _
                          ByVal matrixCode As String, ByVal hoursAgo As Long)
    XOne "INSERT INTO tblSamples (WorkOrderID, ClientSampleID, LocationID, MatrixID, SampleType, CollectedDate, CollectedBy, " & _
         "Containers, Preservative, Status) SELECT " & woID & ", " & Q(clientSampleID) & ", l.LocationID, mx.MatrixID, 'Grab', " & _
         "DateAdd('h', -" & hoursAgo & ", Now()), 'P. Morgan', 3, 'HNO3 / Cool', 'Received' " & _
         "FROM tblLocations AS l, tblMatrices AS mx WHERE l.ProjectID=" & projectID & " AND l.LocationCode=" & Q(locationCode) & _
         " AND mx.MatrixCode=" & Q(matrixCode), "demo sample " & clientSampleID
End Sub

Private Sub SetDemoResult(ByVal labSampleID As String, ByVal analyteName As String, ByVal value As String, ByVal analystID As Long)
    XOne "UPDATE ((tblResults AS r INNER JOIN tblSampleTests AS t ON r.TestID = t.TestID) " & _
         "INNER JOIN tblSamples AS s ON t.SampleID = s.SampleID) " & _
         "INNER JOIN tblAnalytes AS a ON r.AnalyteID = a.AnalyteID " & _
         "SET r.ResultValue = " & value & ", r.AnalysisDate = Now(), r.AnalystID = " & analystID & " " & _
         "WHERE s.LabSampleID=" & Q(labSampleID) & " AND a.AnalyteName=" & Q(analyteName), _
         "demo result " & labSampleID & " / " & analyteName
End Sub

Private Sub AddDemoQC(ByVal batchID As Long, ByVal qcType As String, ByVal analyteName As String, ByVal spike As String, _
                      ByVal background As String, ByVal measured As String, ByVal dup As String, ByVal lo As String, _
                      ByVal hi As String, ByVal rpd As String)
    XOne "INSERT INTO tblQCResults (BatchID, QCType, AnalyteID, SpikeAmount, BackgroundValue, MeasuredValue, DuplicateValue, " & _
         "Units, LowerLimit, UpperLimit, RPDLimit) SELECT " & batchID & ", " & Q(qcType) & ", AnalyteID, " & spike & ", " & _
         background & ", " & measured & ", " & dup & ", 'mg/L', " & lo & ", " & hi & ", " & rpd & " " & _
         "FROM tblAnalytes WHERE AnalyteName=" & Q(analyteName), "demo QC " & qcType
End Sub

Public Sub SeedDemoData()
    Dim db As DAO.Database, clientID As Long, projectID As Long, woID As Long, woNum As String
    Dim a1 As Long, a2 As Long, m2008 As Long, m3000 As Long, icpms As Long, batchID As Long, testID As Long
    Set db = CurrentDb

    X "INSERT INTO tblAnalysts (FirstName, LastName, Initials, Title, Active) VALUES ('Alex', 'Rivera', 'AR', 'Chemist', True)"
    a1 = DMax("AnalystID", "tblAnalysts")
    X "INSERT INTO tblAnalysts (FirstName, LastName, Initials, Title, Active) VALUES ('Jordan', 'Lee', 'JL', 'QA Manager', True)"
    a2 = DMax("AnalystID", "tblAnalysts")

    X "INSERT INTO tblInstruments (InstrumentName, Model, SerialNumber, LastCalibration, NextCalibrationDue, Active) " & _
      "VALUES ('ICP-MS #1', 'ICP-MS', 'SN-1001', DateAdd('d', -20, Date()), DateAdd('d', 10, Date()), True)"
    icpms = DMax("InstrumentID", "tblInstruments")
    X "INSERT INTO tblInstruments (InstrumentName, Model, SerialNumber, LastCalibration, NextCalibrationDue, Active) " & _
      "VALUES ('IC #1', 'Ion Chromatograph', 'SN-2001', DateAdd('d', -60, Date()), DateAdd('d', 120, Date()), True)"
    X "INSERT INTO tblInstruments (InstrumentName, Model, SerialNumber, LastCalibration, NextCalibrationDue, Active) " & _
      "VALUES ('GC-MS #1', 'GC/MS', 'SN-3001', DateAdd('d', -200, Date()), DateAdd('d', -5, Date()), True)"

    X "INSERT INTO tblClients (ClientName, ContactName, Phone, Email, Address, City, State, PostalCode, Active) " & _
      "VALUES ('Example Water District', 'Pat Morgan', '(555) 555-0142', 'lab-contact@example.com', '100 Main Street', " & _
      "'Springfield', 'ST', '00000', True)"
    clientID = DMax("ClientID", "tblClients")

    X "INSERT INTO tblProjects (ClientID, ProjectName, ProjectCode, SiteAddress, LimitSet, PermitNumber, StartDate, Active) " & _
      "VALUES (" & clientID & ", 'Quarterly Compliance Monitoring', 'QCM', 'District service area', 'EPA Primary MCL', " & _
      "'PWS-0000000', Date(), True)"
    projectID = DMax("ProjectID", "tblProjects")

    X "INSERT INTO tblLocations (ProjectID, LocationCode, Description, LocationType, Active) VALUES (" & projectID & _
      ", 'EP-001', 'Treatment plant entry point', 'Entry Point', True)"
    X "INSERT INTO tblLocations (ProjectID, LocationCode, Description, LocationType, Active) VALUES (" & projectID & _
      ", 'DS-101', 'Distribution system tap - 4th Street', 'Distribution Tap', True)"
    X "INSERT INTO tblLocations (ProjectID, LocationCode, Description, LocationType, Active) VALUES (" & projectID & _
      ", 'MW-01', 'Wellfield monitoring well', 'Monitoring Well', True)"

    woNum = NextWorkOrderNumber()
    X "INSERT INTO tblWorkOrders (WorkOrderNumber, ClientID, ProjectID, ReceivedDate, ReceivedByID, DueDate, Status, PONumber, " & _
      "CoolerTempC, ReceivedOnIce, CustodySealsIntact) VALUES (" & Q(woNum) & ", " & clientID & ", " & projectID & ", " & _
      "DateAdd('h', -20, Now()), " & a1 & ", DateAdd('d', 10, Date()), 'Received', 'PO-1001', 3.8, True, True)"
    woID = DMax("WorkOrderID", "tblWorkOrders")

    AddDemoSample woID, projectID, "EP-001-Q4", "EP-001", "DW", 26
    AddDemoSample woID, projectID, "DS-101-Q4", "DS-101", "DW", 25
    AddDemoSample woID, projectID, "MW-01-Q4", "MW-01", "GW", 24
    AssignLabSampleIDs woID, True

    m2008 = DLookup("MethodID", "tblMethods", "MethodCode='EPA 200.8'")
    m3000 = DLookup("MethodID", "tblMethods", "MethodCode='EPA 300.0'")
    AddMethodToWorkOrder woID, m2008, True
    AddMethodToWorkOrder woID, m3000, True

    X "INSERT INTO tblCustodyLog (WorkOrderID, EventDate, CustodyAction, RelinquishedBy, ReceivedBy, StorageLocation, Comments) " & _
      "VALUES (" & woID & ", DateAdd('h', -21, Now()), 'Relinquished by Sampler', 'P. Morgan', 'Courier', 'Cooler 1', 'Seals intact')"
    X "INSERT INTO tblCustodyLog (WorkOrderID, EventDate, CustodyAction, RelinquishedBy, ReceivedBy, StorageLocation, Comments) " & _
      "VALUES (" & woID & ", DateAdd('h', -20, Now()), 'Received by Lab', 'Courier', 'A. Rivera', 'Walk-in cooler', '3.8 C on ice')"

    ' Metals results (one lead result above the action level, one arsenic above the MCL)
    SetDemoResult woNum & "-01", "Arsenic", "0.0021", a1
    SetDemoResult woNum & "-01", "Cadmium", "0.00005", a1
    SetDemoResult woNum & "-01", "Chromium", "0.0035", a1
    SetDemoResult woNum & "-01", "Copper", "0.21", a1
    SetDemoResult woNum & "-01", "Lead", "0.0031", a1
    SetDemoResult woNum & "-01", "Selenium", "0.0008", a1
    SetDemoResult woNum & "-02", "Arsenic", "0.0018", a1
    SetDemoResult woNum & "-02", "Cadmium", "0.0002", a1
    SetDemoResult woNum & "-02", "Chromium", "0.0029", a1
    SetDemoResult woNum & "-02", "Copper", "0.48", a1
    SetDemoResult woNum & "-02", "Lead", "0.022", a1
    SetDemoResult woNum & "-02", "Selenium", "0.0006", a1
    SetDemoResult woNum & "-03", "Arsenic", "0.014", a1
    SetDemoResult woNum & "-03", "Lead", "0.0009", a1
    ' Anions on the first sample
    SetDemoResult woNum & "-01", "Nitrate as N", "2.4", a1
    SetDemoResult woNum & "-01", "Chloride", "38", a1
    SetDemoResult woNum & "-01", "Sulfate", "41", a1
    SetDemoResult woNum & "-01", "Fluoride", "0.7", a1

    X "UPDATE tblResults SET Qualifier = IIf(ResultValue < MDL, 'U', IIf(ResultValue < RL, 'J', Null)) WHERE ResultValue Is Not Null"
    X "UPDATE tblSampleTests SET Status = 'In Progress', AnalystID = " & a1 & ", AnalysisDate = Now(), PrepDate = DateAdd('h', -4, Now()) " & _
      "WHERE TestID IN (SELECT TestID FROM tblResults WHERE ResultValue Is Not Null)"

    ' QC batch for the metals run (the MSD fails recovery on purpose)
    X "INSERT INTO tblQCBatches (BatchNumber, MethodID, InstrumentID, AnalystID, BatchDate, Status) VALUES ('B" & _
      Format$(Date, "yymmdd") & "-01', " & m2008 & ", " & icpms & ", " & a1 & ", Now(), 'Open')"
    batchID = DMax("BatchID", "tblQCBatches")
    X "UPDATE tblSampleTests SET BatchID = " & batchID & " WHERE MethodID = " & m2008
    AddDemoQC batchID, "Method Blank", "Lead", "Null", "Null", "0.0001", "Null", "Null", "0.001", "Null"
    AddDemoQC batchID, "LCS", "Lead", "0.05", "Null", "0.0512", "Null", "85", "115", "Null"
    AddDemoQC batchID, "MS", "Lead", "0.05", "0.0031", "0.0499", "Null", "85", "115", "Null"
    AddDemoQC batchID, "MSD", "Lead", "0.05", "0.0031", "0.0405", "0.0499", "85", "115", "20"

    ' Review the metals test on sample -01
    testID = DLookup("TestID", "tblSampleTests", "MethodID=" & m2008 & " AND SampleID=" & _
             DLookup("SampleID", "tblSamples", "LabSampleID=" & Q(woNum & "-01")))
    db.Execute "UPDATE tblResults SET Reviewed=True, ReviewedBy='JL', ReviewedDate=Now() WHERE TestID=" & testID, dbFailOnError
    db.Execute "UPDATE tblSampleTests SET Status='Reviewed' WHERE TestID=" & testID, dbFailOnError
    UpdateWorkOrderStatus woID
    LogAudit "tblWorkOrders", woID, "Demo data loaded"
End Sub
