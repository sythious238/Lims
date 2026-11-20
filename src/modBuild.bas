Attribute VB_Name = "modBuild"
Option Compare Database
Option Explicit

' =============================================================================
'  EnviroLIMS - database builder
'
'  BuildLIMS            Creates tables, relationships, reference data, queries,
'                       forms and reports in the current (blank) database.
'                       BuildLIMS False        -> no demo data
'                       BuildLIMS True, True   -> DROP everything and rebuild
'  RebuildUI            Regenerates only queries, forms and reports. Data is kept.
' =============================================================================

Private db As DAO.Database
Private m_step As String

Private Const LIMS_TABLES As String = _
    "tblAuditLog,tblCustodyLog,tblQCResults,tblResults,tblSampleTests,tblQCBatches," & _
    "tblSamples,tblWorkOrders,tblRegulatoryLimits,tblMethodAnalytes,tblMethods,tblAnalytes," & _
    "tblMatrices,tblLocations,tblProjects,tblClients,tblAnalysts,tblInstruments,tblSettings"

Public Sub BuildLIMS(Optional ByVal includeDemoData As Boolean = True, Optional ByVal forceRebuild As Boolean = False)
    On Error GoTo EH
    Set db = CurrentDb

    If TableExists("tblSamples") And Not forceRebuild Then
        MsgBox "EnviroLIMS tables already exist in this database." & vbCrLf & vbCrLf & _
               "- Run RebuildUI to regenerate queries, forms and reports (data is kept)." & vbCrLf & _
               "- Run BuildLIMS True, True to DELETE ALL DATA and rebuild from scratch.", _
               vbExclamation, "EnviroLIMS"
        Exit Sub
    End If

    DoCmd.Hourglass True
    Stage "Removing old objects":   DropSchema
    Stage "Creating tables":        CreateTables
    Stage "Creating indexes":       CreateIndexes
    Stage "Setting field defaults": SetFieldProperties
    Stage "Creating relationships": CreateRelationships
    Stage "Loading reference data": SeedReferenceData
    Stage "Creating queries":       CreateQueries
    Stage "Creating forms and reports": BuildAllUI
    If includeDemoData Then
        Stage "Loading demo data":  SeedDemoData
    End If
    Stage "Setting startup options": SetStartupOptions

    SysCmd acSysCmdClearStatus
    DoCmd.Hourglass False
    MsgBox "EnviroLIMS was built successfully." & vbCrLf & vbCrLf & _
           "Next steps:" & vbCrLf & _
           "  1. Setup > Settings: enter your laboratory name and address." & vbCrLf & _
           "  2. Setup > Methods: replace the example MDLs/RLs with your own." & vbCrLf & _
           "  3. Setup > Regulatory Limits: confirm limits for your jurisdiction.", _
           vbInformation, "EnviroLIMS"
    DoCmd.OpenForm "frmMainMenu"
    Exit Sub
EH:
    Dim msg As String
    msg = "Build failed while: " & m_step & vbCrLf & vbCrLf & "Error " & Err.Number & ": " & Err.Description
    Application.Echo True
    DoCmd.Hourglass False
    SysCmd acSysCmdClearStatus
    MsgBox msg, vbCritical, "EnviroLIMS"
End Sub

Public Sub RebuildUI()
    On Error GoTo EH
    Set db = CurrentDb
    DoCmd.Hourglass True
    Stage "Creating queries":           CreateQueries
    Stage "Creating forms and reports": BuildAllUI
    SysCmd acSysCmdClearStatus
    DoCmd.Hourglass False
    MsgBox "Queries, forms and reports were regenerated.", vbInformation, "EnviroLIMS"
    Exit Sub
EH:
    Dim msg As String
    msg = "Rebuild failed while: " & m_step & vbCrLf & vbCrLf & "Error " & Err.Number & ": " & Err.Description
    Application.Echo True
    DoCmd.Hourglass False
    SysCmd acSysCmdClearStatus
    MsgBox msg, vbCritical, "EnviroLIMS"
End Sub

' ----------------------------------------------------------------------------
'  Helpers
' ----------------------------------------------------------------------------

Private Sub Stage(ByVal stepName As String)
    m_step = stepName
    SysCmd acSysCmdSetStatus, "EnviroLIMS: " & stepName & "..."
    DoEvents
End Sub

Private Sub X(ByVal sql As String)
    db.Execute sql, dbFailOnError
End Sub

Public Function TableExists(ByVal tableName As String) As Boolean
    Dim td As DAO.TableDef
    For Each td In CurrentDb.TableDefs
        If StrComp(td.Name, tableName, vbTextCompare) = 0 Then
            TableExists = True
            Exit Function
        End If
    Next td
End Function

Private Function IsLimsTable(ByVal tableName As String) As Boolean
    IsLimsTable = InStr(1, "," & LIMS_TABLES & ",", "," & tableName & ",", vbTextCompare) > 0
End Function

' ----------------------------------------------------------------------------
'  Schema
' ----------------------------------------------------------------------------

Private Sub DropSchema()
    Dim i As Long, rel As DAO.Relation, names() As String
    For i = db.Relations.Count - 1 To 0 Step -1
        Set rel = db.Relations(i)
        If IsLimsTable(rel.Table) And IsLimsTable(rel.ForeignTable) Then db.Relations.Delete rel.Name
    Next i
    names = Split(LIMS_TABLES, ",")
    For i = LBound(names) To UBound(names)
        If TableExists(names(i)) Then X "DROP TABLE [" & names(i) & "]"
    Next i
    db.TableDefs.Refresh
End Sub

Private Sub CreateTables()
    X "CREATE TABLE tblSettings (SettingName TEXT(50) CONSTRAINT pk_tblSettings PRIMARY KEY, SettingValue TEXT(255), Description TEXT(255))"

    X "CREATE TABLE tblClients (ClientID COUNTER CONSTRAINT pk_tblClients PRIMARY KEY, ClientName TEXT(100) NOT NULL, " & _
      "ContactName TEXT(100), Phone TEXT(30), Email TEXT(100), Address TEXT(255), City TEXT(50), State TEXT(20), " & _
      "PostalCode TEXT(15), Active YESNO, Notes MEMO)"

    X "CREATE TABLE tblProjects (ProjectID COUNTER CONSTRAINT pk_tblProjects PRIMARY KEY, ClientID LONG NOT NULL, " & _
      "ProjectName TEXT(100) NOT NULL, ProjectCode TEXT(30), SiteAddress TEXT(255), LimitSet TEXT(50), " & _
      "PermitNumber TEXT(50), StartDate DATETIME, Active YESNO, Notes MEMO)"

    X "CREATE TABLE tblLocations (LocationID COUNTER CONSTRAINT pk_tblLocations PRIMARY KEY, ProjectID LONG NOT NULL, " & _
      "LocationCode TEXT(30) NOT NULL, Description TEXT(255), LocationType TEXT(50), Latitude DOUBLE, Longitude DOUBLE, Active YESNO)"

    X "CREATE TABLE tblMatrices (MatrixID COUNTER CONSTRAINT pk_tblMatrices PRIMARY KEY, MatrixCode TEXT(10) NOT NULL, " & _
      "MatrixName TEXT(50) NOT NULL, MatrixClass TEXT(20))"

    X "CREATE TABLE tblAnalytes (AnalyteID COUNTER CONSTRAINT pk_tblAnalytes PRIMARY KEY, AnalyteName TEXT(100) NOT NULL, " & _
      "CASNumber TEXT(20), AnalyteGroup TEXT(50))"

    X "CREATE TABLE tblMethods (MethodID COUNTER CONSTRAINT pk_tblMethods PRIMARY KEY, MethodCode TEXT(50) NOT NULL, " & _
      "MethodName TEXT(150), MatrixClass TEXT(20), HoldTimeHours DOUBLE, Preservation TEXT(150), ContainerType TEXT(100), " & _
      "TurnaroundDays LONG, Active YESNO)"

    X "CREATE TABLE tblMethodAnalytes (MethodAnalyteID COUNTER CONSTRAINT pk_tblMethodAnalytes PRIMARY KEY, " & _
      "MethodID LONG NOT NULL, AnalyteID LONG NOT NULL, Units TEXT(20), MDL DOUBLE, RL DOUBLE, " & _
      "RecoveryLow DOUBLE, RecoveryHigh DOUBLE, RPDLimit DOUBLE)"

    X "CREATE TABLE tblRegulatoryLimits (LimitID COUNTER CONSTRAINT pk_tblRegulatoryLimits PRIMARY KEY, " & _
      "LimitSet TEXT(50) NOT NULL, AnalyteID LONG NOT NULL, MatrixID LONG NOT NULL, LimitValue DOUBLE, " & _
      "Units TEXT(20), LimitType TEXT(50), Source TEXT(150))"

    X "CREATE TABLE tblAnalysts (AnalystID COUNTER CONSTRAINT pk_tblAnalysts PRIMARY KEY, FirstName TEXT(50), " & _
      "LastName TEXT(50), Initials TEXT(5) NOT NULL, Title TEXT(50), Active YESNO)"

    X "CREATE TABLE tblInstruments (InstrumentID COUNTER CONSTRAINT pk_tblInstruments PRIMARY KEY, " & _
      "InstrumentName TEXT(100) NOT NULL, Model TEXT(100), SerialNumber TEXT(50), LastCalibration DATETIME, " & _
      "NextCalibrationDue DATETIME, Active YESNO)"

    X "CREATE TABLE tblWorkOrders (WorkOrderID COUNTER CONSTRAINT pk_tblWorkOrders PRIMARY KEY, " & _
      "WorkOrderNumber TEXT(20) NOT NULL, ClientID LONG NOT NULL, ProjectID LONG, ReceivedDate DATETIME, " & _
      "ReceivedByID LONG, DueDate DATETIME, Status TEXT(20), PONumber TEXT(50), CoolerTempC DOUBLE, " & _
      "ReceivedOnIce YESNO, CustodySealsIntact YESNO, Notes MEMO)"

    X "CREATE TABLE tblSamples (SampleID COUNTER CONSTRAINT pk_tblSamples PRIMARY KEY, WorkOrderID LONG NOT NULL, " & _
      "LabSampleID TEXT(30), ClientSampleID TEXT(50), LocationID LONG, MatrixID LONG NOT NULL, SampleType TEXT(30), " & _
      "CollectedDate DATETIME, CollectedBy TEXT(100), Containers LONG, Preservative TEXT(50), FieldPH DOUBLE, " & _
      "FieldTempC DOUBLE, Status TEXT(20), DisposalDate DATETIME, Notes MEMO)"

    X "CREATE TABLE tblQCBatches (BatchID COUNTER CONSTRAINT pk_tblQCBatches PRIMARY KEY, BatchNumber TEXT(30) NOT NULL, " & _
      "MethodID LONG NOT NULL, InstrumentID LONG, AnalystID LONG, BatchDate DATETIME, Status TEXT(20), Notes MEMO)"

    X "CREATE TABLE tblSampleTests (TestID COUNTER CONSTRAINT pk_tblSampleTests PRIMARY KEY, SampleID LONG NOT NULL, " & _
      "MethodID LONG NOT NULL, Status TEXT(20), AnalystID LONG, BatchID LONG, PrepDate DATETIME, AnalysisDate DATETIME, " & _
      "DueDate DATETIME, Notes MEMO)"

    X "CREATE TABLE tblResults (ResultID COUNTER CONSTRAINT pk_tblResults PRIMARY KEY, TestID LONG NOT NULL, " & _
      "AnalyteID LONG NOT NULL, ResultValue DOUBLE, Qualifier TEXT(10), Units TEXT(20), DilutionFactor DOUBLE, " & _
      "MDL DOUBLE, RL DOUBLE, AnalystID LONG, AnalysisDate DATETIME, Reviewed YESNO, ReviewedBy TEXT(50), " & _
      "ReviewedDate DATETIME, Comments TEXT(255))"

    X "CREATE TABLE tblQCResults (QCResultID COUNTER CONSTRAINT pk_tblQCResults PRIMARY KEY, BatchID LONG NOT NULL, " & _
      "QCType TEXT(30) NOT NULL, AnalyteID LONG NOT NULL, ParentSampleID LONG, SpikeAmount DOUBLE, BackgroundValue DOUBLE, " & _
      "MeasuredValue DOUBLE, DuplicateValue DOUBLE, Units TEXT(20), LowerLimit DOUBLE, UpperLimit DOUBLE, " & _
      "RPDLimit DOUBLE, Comments TEXT(255))"

    X "CREATE TABLE tblCustodyLog (CustodyID COUNTER CONSTRAINT pk_tblCustodyLog PRIMARY KEY, WorkOrderID LONG NOT NULL, " & _
      "EventDate DATETIME, CustodyAction TEXT(50), RelinquishedBy TEXT(100), ReceivedBy TEXT(100), " & _
      "StorageLocation TEXT(100), Comments TEXT(255))"

    X "CREATE TABLE tblAuditLog (AuditID COUNTER CONSTRAINT pk_tblAuditLog PRIMARY KEY, EventTime DATETIME, " & _
      "UserName TEXT(50), TableName TEXT(50), RecordID LONG, ActionTaken TEXT(50), Details TEXT(255))"

    db.TableDefs.Refresh
End Sub

Private Sub CreateIndexes()
    X "CREATE UNIQUE INDEX ux_WorkOrderNumber ON tblWorkOrders (WorkOrderNumber)"
    X "CREATE UNIQUE INDEX ux_LabSampleID ON tblSamples (LabSampleID) WITH IGNORE NULL"
    X "CREATE UNIQUE INDEX ux_MethodAnalyte ON tblMethodAnalytes (MethodID, AnalyteID)"
    X "CREATE UNIQUE INDEX ux_ResultTestAnalyte ON tblResults (TestID, AnalyteID)"
    X "CREATE UNIQUE INDEX ux_Limit ON tblRegulatoryLimits (LimitSet, AnalyteID, MatrixID)"
    X "CREATE UNIQUE INDEX ux_BatchNumber ON tblQCBatches (BatchNumber)"
    X "CREATE UNIQUE INDEX ux_MatrixCode ON tblMatrices (MatrixCode)"
    X "CREATE UNIQUE INDEX ux_MethodCode ON tblMethods (MethodCode)"
    X "CREATE UNIQUE INDEX ux_AnalyteName ON tblAnalytes (AnalyteName)"
    X "CREATE UNIQUE INDEX ux_LocationCode ON tblLocations (ProjectID, LocationCode)"
    X "CREATE INDEX ix_TestStatus ON tblSampleTests (Status)"
    X "CREATE INDEX ix_TestDue ON tblSampleTests (DueDate)"
    X "CREATE INDEX ix_CollectedDate ON tblSamples (CollectedDate)"
    X "CREATE INDEX ix_AuditTime ON tblAuditLog (EventTime)"
End Sub

Private Sub SetDefault(ByVal tableName As String, ByVal fieldName As String, ByVal defaultValue As String)
    db.TableDefs(tableName).Fields(fieldName).DefaultValue = defaultValue
End Sub

Private Sub SetRule(ByVal tableName As String, ByVal fieldName As String, ByVal rule As String, ByVal ruleText As String)
    With db.TableDefs(tableName).Fields(fieldName)
        .ValidationRule = rule
        .ValidationText = ruleText
    End With
End Sub

' Shows Yes/No fields as check boxes in table datasheets.
Private Sub YesNoAsCheckbox()
    Dim td As DAO.TableDef, fld As DAO.Field
    On Error Resume Next
    For Each td In db.TableDefs
        If IsLimsTable(td.Name) Then
            For Each fld In td.Fields
                If fld.Type = dbBoolean Then
                    fld.Properties("DisplayControl") = acCheckBox
                    If Err.Number <> 0 Then
                        Err.Clear
                        fld.Properties.Append fld.CreateProperty("DisplayControl", dbInteger, acCheckBox)
                    End If
                    Err.Clear
                End If
            Next fld
        End If
    Next td
End Sub

Private Sub SetFieldProperties()
    SetDefault "tblClients", "Active", "True"
    SetDefault "tblProjects", "Active", "True"
    SetDefault "tblLocations", "Active", "True"
    SetDefault "tblMethods", "Active", "True"
    SetDefault "tblMethods", "TurnaroundDays", "10"
    SetDefault "tblAnalysts", "Active", "True"
    SetDefault "tblInstruments", "Active", "True"
    SetDefault "tblWorkOrders", "Status", """Received"""
    SetDefault "tblWorkOrders", "ReceivedDate", "Now()"
    SetDefault "tblWorkOrders", "ReceivedOnIce", "False"
    SetDefault "tblWorkOrders", "CustodySealsIntact", "False"
    SetDefault "tblSamples", "Status", """Received"""
    SetDefault "tblSamples", "SampleType", """Grab"""
    SetDefault "tblSamples", "Containers", "1"
    SetDefault "tblSampleTests", "Status", """Pending"""
    SetDefault "tblResults", "DilutionFactor", "1"
    SetDefault "tblResults", "Reviewed", "False"
    SetDefault "tblQCBatches", "Status", """Open"""
    SetDefault "tblQCBatches", "BatchDate", "Now()"
    SetDefault "tblCustodyLog", "EventDate", "Now()"
    SetDefault "tblAuditLog", "EventTime", "Now()"

    SetRule "tblResults", "DilutionFactor", ">0 Or Is Null", "Dilution factor must be greater than zero."
    SetRule "tblResults", "MDL", ">=0 Or Is Null", "MDL cannot be negative."
    SetRule "tblResults", "RL", ">=0 Or Is Null", "RL cannot be negative."
    SetRule "tblSamples", "FieldPH", "Between 0 And 14 Or Is Null", "pH must be between 0 and 14."
    SetRule "tblMethods", "HoldTimeHours", ">0 Or Is Null", "Hold time must be greater than zero."

    YesNoAsCheckbox
End Sub

Private Sub AddRel(ByVal pkTable As String, ByVal pkField As String, ByVal fkTable As String, ByVal fkField As String, _
                   Optional ByVal cascadeDelete As Boolean = False)
    Dim rel As DAO.Relation, fld As DAO.Field
    Set rel = db.CreateRelation("rel_" & fkTable & "_" & fkField, pkTable, fkTable, IIf(cascadeDelete, dbRelationDeleteCascade, 0))
    Set fld = rel.CreateField(pkField)
    fld.ForeignName = fkField
    rel.Fields.Append fld
    db.Relations.Append rel
End Sub

Private Sub CreateRelationships()
    AddRel "tblClients", "ClientID", "tblProjects", "ClientID"
    AddRel "tblProjects", "ProjectID", "tblLocations", "ProjectID", True
    AddRel "tblMethods", "MethodID", "tblMethodAnalytes", "MethodID", True
    AddRel "tblAnalytes", "AnalyteID", "tblMethodAnalytes", "AnalyteID"
    AddRel "tblAnalytes", "AnalyteID", "tblRegulatoryLimits", "AnalyteID"
    AddRel "tblMatrices", "MatrixID", "tblRegulatoryLimits", "MatrixID"
    AddRel "tblClients", "ClientID", "tblWorkOrders", "ClientID"
    AddRel "tblProjects", "ProjectID", "tblWorkOrders", "ProjectID"
    AddRel "tblAnalysts", "AnalystID", "tblWorkOrders", "ReceivedByID"
    AddRel "tblWorkOrders", "WorkOrderID", "tblSamples", "WorkOrderID", True
    AddRel "tblLocations", "LocationID", "tblSamples", "LocationID"
    AddRel "tblMatrices", "MatrixID", "tblSamples", "MatrixID"
    AddRel "tblSamples", "SampleID", "tblSampleTests", "SampleID", True
    AddRel "tblMethods", "MethodID", "tblSampleTests", "MethodID"
    AddRel "tblAnalysts", "AnalystID", "tblSampleTests", "AnalystID"
    AddRel "tblQCBatches", "BatchID", "tblSampleTests", "BatchID"
    AddRel "tblSampleTests", "TestID", "tblResults", "TestID", True
    AddRel "tblAnalytes", "AnalyteID", "tblResults", "AnalyteID"
    AddRel "tblAnalysts", "AnalystID", "tblResults", "AnalystID"
    AddRel "tblMethods", "MethodID", "tblQCBatches", "MethodID"
    AddRel "tblInstruments", "InstrumentID", "tblQCBatches", "InstrumentID"
    AddRel "tblAnalysts", "AnalystID", "tblQCBatches", "AnalystID"
    AddRel "tblQCBatches", "BatchID", "tblQCResults", "BatchID", True
    AddRel "tblAnalytes", "AnalyteID", "tblQCResults", "AnalyteID"
    AddRel "tblSamples", "SampleID", "tblQCResults", "ParentSampleID"
    AddRel "tblWorkOrders", "WorkOrderID", "tblCustodyLog", "WorkOrderID", True
End Sub

' ----------------------------------------------------------------------------
'  Queries
' ----------------------------------------------------------------------------

Private Sub CreateQ(ByVal queryName As String, ByVal sql As String)
    On Error Resume Next
    db.QueryDefs.Delete queryName
    On Error GoTo 0
    db.CreateQueryDef queryName, sql
End Sub

Private Sub CreateQueries()
    Dim s As String

    ' --- Every result with its sample, work order, client, method, location ---
    s = "SELECT r.ResultID, r.TestID, t.SampleID, s.WorkOrderID, w.WorkOrderNumber, w.ClientID, c.ClientName, "
    s = s & "w.ProjectID, p.ProjectName, p.LimitSet, w.ReceivedDate, s.LabSampleID, s.ClientSampleID, "
    s = s & "s.LocationID, l.LocationCode, s.MatrixID, mx.MatrixName, s.SampleType, s.CollectedDate, "
    s = s & "t.MethodID, mt.MethodCode, t.Status AS TestStatus, r.AnalyteID, a.AnalyteName, a.CASNumber, "
    s = s & "r.ResultValue, r.Qualifier, r.Units, r.DilutionFactor, r.MDL, r.RL, r.AnalysisDate, "
    s = s & "anl.Initials AS AnalystInitials, r.Reviewed, r.ReviewedBy, r.ReviewedDate, r.Comments, "
    s = s & "ReportValue(r.ResultValue, r.Qualifier) AS ReportedResult "
    s = s & "FROM (((((((((tblResults AS r "
    s = s & "INNER JOIN tblSampleTests AS t ON r.TestID = t.TestID) "
    s = s & "INNER JOIN tblSamples AS s ON t.SampleID = s.SampleID) "
    s = s & "INNER JOIN tblWorkOrders AS w ON s.WorkOrderID = w.WorkOrderID) "
    s = s & "INNER JOIN tblClients AS c ON w.ClientID = c.ClientID) "
    s = s & "INNER JOIN tblMatrices AS mx ON s.MatrixID = mx.MatrixID) "
    s = s & "INNER JOIN tblMethods AS mt ON t.MethodID = mt.MethodID) "
    s = s & "INNER JOIN tblAnalytes AS a ON r.AnalyteID = a.AnalyteID) "
    s = s & "LEFT JOIN tblProjects AS p ON w.ProjectID = p.ProjectID) "
    s = s & "LEFT JOIN tblLocations AS l ON s.LocationID = l.LocationID) "
    s = s & "LEFT JOIN tblAnalysts AS anl ON r.AnalystID = anl.AnalystID"
    CreateQ "qryResultDetail", s

    ' --- Results compared against the project's regulatory limit set ---
    s = "SELECT d.*, lim.LimitValue, lim.LimitType, lim.Units AS LimitUnits, "
    s = s & "IIf(lim.LimitValue Is Null Or d.ResultValue Is Null, '', "
    s = s & "IIf(Nz(d.Qualifier,'') Like '*U*', '', IIf(d.ResultValue > lim.LimitValue, 'EXCEEDS', ''))) AS ExceedFlag "
    s = s & "FROM qryResultDetail AS d LEFT JOIN tblRegulatoryLimits AS lim "
    s = s & "ON (d.AnalyteID = lim.AnalyteID) AND (d.MatrixID = lim.MatrixID) AND (d.LimitSet = lim.LimitSet)"
    CreateQ "qryResultsWithLimits", s

    CreateQ "qryExceedances", "SELECT * FROM qryResultsWithLimits WHERE ExceedFlag = 'EXCEEDS' ORDER BY CollectedDate DESC"

    CreateQ "qryUnreviewedResults", "SELECT * FROM qryResultDetail WHERE Reviewed = False AND ResultValue Is Not Null " & _
            "ORDER BY WorkOrderNumber, LabSampleID, MethodCode, AnalyteName"

    s = "SELECT ClientName, ProjectName, LocationCode, MatrixName, AnalyteName, CollectedDate, ResultValue, "
    s = s & "Qualifier, Units, LimitValue, ExceedFlag, LabSampleID FROM qryResultsWithLimits "
    s = s & "WHERE ResultValue Is Not Null ORDER BY LocationCode, AnalyteName, CollectedDate"
    CreateQ "qryLocationTrends", s

    ' --- Holding times ---
    s = "SELECT t.TestID, t.SampleID, s.WorkOrderID, w.WorkOrderNumber, c.ClientName, s.LabSampleID, "
    s = s & "s.ClientSampleID, t.MethodID, mt.MethodCode, t.Status, t.AnalystID, t.DueDate, t.AnalysisDate, "
    s = s & "s.CollectedDate, mt.HoldTimeHours, "
    s = s & "IIf(s.CollectedDate Is Null Or mt.HoldTimeHours Is Null, Null, "
    s = s & "DateAdd('n', mt.HoldTimeHours * 60, s.CollectedDate)) AS HoldExpires "
    s = s & "FROM (((tblSampleTests AS t "
    s = s & "INNER JOIN tblSamples AS s ON t.SampleID = s.SampleID) "
    s = s & "INNER JOIN tblWorkOrders AS w ON s.WorkOrderID = w.WorkOrderID) "
    s = s & "INNER JOIN tblClients AS c ON w.ClientID = c.ClientID) "
    s = s & "INNER JOIN tblMethods AS mt ON t.MethodID = mt.MethodID"
    CreateQ "qryHoldTimeBase", s

    s = "SELECT b.*, "
    s = s & "IIf(b.HoldExpires Is Null, 'N/A', IIf(b.AnalysisDate Is Null, "
    s = s & "IIf(Now() > b.HoldExpires, 'EXPIRED', IIf(DateAdd('h', 24, Now()) > b.HoldExpires, 'DUE <24H', 'OK')), "
    s = s & "IIf(b.AnalysisDate > b.HoldExpires, 'OUT OF HOLD', 'MET'))) AS HoldStatus, "
    s = s & "IIf(b.HoldExpires Is Null Or b.AnalysisDate Is Not Null, Null, Round((b.HoldExpires - Now()) * 24, 1)) AS HoursRemaining "
    s = s & "FROM qryHoldTimeBase AS b"
    CreateQ "qryHoldTimeStatus", s

    CreateQ "qryOpenTests", "SELECT * FROM qryHoldTimeStatus WHERE Nz(Status,'') Not In ('Reviewed','Cancelled') " & _
            "ORDER BY HoldExpires, DueDate"
    CreateQ "qryOverdueTests", "SELECT * FROM qryOpenTests WHERE DueDate < Date() ORDER BY DueDate"

    ' --- Quality control ---
    s = "SELECT q.QCResultID, q.BatchID, b.BatchNumber, b.BatchDate, mt.MethodCode, q.QCType, a.AnalyteName, "
    s = s & "q.SpikeAmount, q.BackgroundValue, q.MeasuredValue, q.DuplicateValue, q.Units, q.LowerLimit, "
    s = s & "q.UpperLimit, q.RPDLimit, CalcRecovery(q.MeasuredValue, q.SpikeAmount, q.BackgroundValue) AS PctRecovery, "
    s = s & "CalcRPD(q.MeasuredValue, q.DuplicateValue) AS RPD, "
    s = s & "EvalQC(q.QCType, q.MeasuredValue, q.SpikeAmount, q.BackgroundValue, q.DuplicateValue, "
    s = s & "q.LowerLimit, q.UpperLimit, q.RPDLimit) AS QCStatus, q.Comments "
    s = s & "FROM ((tblQCResults AS q "
    s = s & "INNER JOIN tblQCBatches AS b ON q.BatchID = b.BatchID) "
    s = s & "INNER JOIN tblMethods AS mt ON b.MethodID = mt.MethodID) "
    s = s & "INNER JOIN tblAnalytes AS a ON q.AnalyteID = a.AnalyteID"
    CreateQ "qryQCEvaluation", s

    CreateQ "qryQCFailures", "SELECT * FROM qryQCEvaluation WHERE QCStatus = 'FAIL' ORDER BY BatchDate DESC"

    ' --- Work order progress ---
    s = "SELECT w.WorkOrderID, w.WorkOrderNumber, c.ClientName, w.ReceivedDate, w.DueDate, w.Status, "
    s = s & "Count(t.TestID) AS TotalTests, Sum(IIf(t.Status = 'Reviewed', 1, 0)) AS ReviewedTests, "
    s = s & "Sum(IIf(t.Status = 'Pending', 1, 0)) AS PendingTests, "
    s = s & "IIf(Count(t.TestID) = 0, 0, Sum(IIf(t.Status = 'Reviewed', 1, 0)) / Count(t.TestID)) AS PctComplete "
    s = s & "FROM ((tblWorkOrders AS w INNER JOIN tblClients AS c ON w.ClientID = c.ClientID) "
    s = s & "LEFT JOIN tblSamples AS s ON w.WorkOrderID = s.WorkOrderID) "
    s = s & "LEFT JOIN tblSampleTests AS t ON s.SampleID = t.SampleID "
    s = s & "GROUP BY w.WorkOrderID, w.WorkOrderNumber, c.ClientName, w.ReceivedDate, w.DueDate, w.Status "
    s = s & "ORDER BY w.WorkOrderNumber DESC"
    CreateQ "qryWorkOrderProgress", s

    ' --- Chain of custody ---
    s = "SELECT cl.CustodyID, cl.WorkOrderID, w.WorkOrderNumber, c.ClientName, cl.EventDate, cl.CustodyAction, "
    s = s & "cl.RelinquishedBy, cl.ReceivedBy, cl.StorageLocation, cl.Comments "
    s = s & "FROM (tblCustodyLog AS cl INNER JOIN tblWorkOrders AS w ON cl.WorkOrderID = w.WorkOrderID) "
    s = s & "INNER JOIN tblClients AS c ON w.ClientID = c.ClientID"
    CreateQ "qryCustodyReport", s

    ' --- Instruments ---
    CreateQ "qryInstrumentCalDue", "SELECT * FROM tblInstruments WHERE Active = True AND " & _
            "(NextCalibrationDue Is Null OR NextCalibrationDue <= DateAdd('d', 30, Date())) ORDER BY NextCalibrationDue"
End Sub

' ----------------------------------------------------------------------------
'  Startup
' ----------------------------------------------------------------------------

Private Sub SetDbProp(ByVal propName As String, ByVal propType As Integer, ByVal propValue As Variant)
    On Error Resume Next
    db.Properties(propName) = propValue
    If Err.Number <> 0 Then
        Err.Clear
        db.Properties.Append db.CreateProperty(propName, propType, propValue)
    End If
End Sub

Private Sub SetStartupOptions()
    SetDbProp "StartupForm", dbText, "frmMainMenu"
    SetDbProp "AppTitle", dbText, "EnviroLIMS"
    On Error Resume Next
    Application.RefreshTitleBar
End Sub
