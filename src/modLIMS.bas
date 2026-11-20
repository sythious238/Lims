Attribute VB_Name = "modLIMS"
Option Compare Database
Option Explicit

' =============================================================================
'  EnviroLIMS - runtime business logic
'
'  Everything in this module is called from form/report event properties
'  (e.g. OnClick = "=AssignLabSampleIDs([WorkOrderID])") or from queries.
'  Keeping the logic in a standard module (instead of form class modules)
'  means the forms can be regenerated at any time with RebuildUI.
' =============================================================================

' ----------------------------------------------------------------------------
'  Small helpers
' ----------------------------------------------------------------------------

' Returns a value quoted for use in Jet SQL ('O''Brien'), or Null.
Public Function SqlText(ByVal v As Variant) As String
    If IsNull(v) Then
        SqlText = "Null"
    Else
        SqlText = "'" & Replace(CStr(v), "'", "''") & "'"
    End If
End Function

' Reads a value from tblSettings.
Public Function LimsSetting(ByVal settingName As String, Optional ByVal defaultValue As String = "") As String
    Dim v As Variant
    On Error Resume Next
    v = DLookup("SettingValue", "tblSettings", "SettingName=" & SqlText(settingName))
    If IsNull(v) Or IsEmpty(v) Then
        LimsSetting = defaultValue
    Else
        LimsSetting = CStr(v)
    End If
End Function

Public Function CurrentLimsUser() As String
    CurrentLimsUser = Environ$("USERNAME")
    If Len(CurrentLimsUser) = 0 Then CurrentLimsUser = "unknown"
End Function

Public Sub LogAudit(ByVal tableName As String, ByVal recordID As Variant, ByVal actionTaken As String, Optional ByVal details As String = "")
    On Error Resume Next
    CurrentDb.Execute "INSERT INTO tblAuditLog (EventTime, UserName, TableName, RecordID, ActionTaken, Details) VALUES (Now(), " & _
        SqlText(CurrentLimsUser()) & ", " & SqlText(tableName) & ", " & IIf(IsNull(recordID), "Null", CLng(Nz(recordID, 0))) & ", " & _
        SqlText(actionTaken) & ", " & SqlText(Left$(details, 255)) & ")", dbFailOnError
End Sub

' Saves the record on the active form, if it has unsaved edits.
Public Sub SaveActiveRecord()
    On Error Resume Next
    If Screen.ActiveForm.Dirty Then Screen.ActiveForm.Dirty = False
End Sub

' Refreshes the active form and requeries all of its subforms
' (without moving the main form off the current record).
Public Sub RequeryActive()
    Dim f As Form, ctl As Control
    On Error Resume Next
    Set f = Screen.ActiveForm
    If f Is Nothing Then Exit Sub
    For Each ctl In f.Controls
        If ctl.ControlType = acSubform Then ctl.Form.Requery
    Next ctl
    f.Refresh
End Sub

' Formats a number to N significant figures (environmental reporting convention).
Public Function FormatSig(ByVal v As Variant, Optional ByVal sigFigs As Integer = 3) As String
    Dim x As Double, d As Integer, factor As Double
    If IsNull(v) Then Exit Function
    x = CDbl(v)
    If x = 0 Then
        FormatSig = "0"
        Exit Function
    End If
    d = sigFigs - 1 - Int(Log(Abs(x)) / Log(10#) + 0.000000001)
    If d > 0 Then
        FormatSig = Format$(Round(x, d), "0." & String$(d, "0"))
    Else
        factor = 10 ^ (-d)
        FormatSig = Format$(Round(x / factor, 0) * factor, "0")
    End If
End Function

' The value printed on reports: "Pending", "ND" or the result to 3 sig figs.
Public Function ReportValue(ByVal v As Variant, ByVal qualifier As Variant) As String
    If IsNull(v) Then
        ReportValue = "Pending"
    ElseIf InStr(1, UCase$(Nz(qualifier, "")), "U") > 0 Then
        ReportValue = "ND"
    Else
        ReportValue = FormatSig(v, 3)
    End If
End Function

' ----------------------------------------------------------------------------
'  Navigation (used by buttons)
' ----------------------------------------------------------------------------

Public Function OpenFormByName(ByVal formName As String) As Boolean
    SaveActiveRecord
    DoCmd.OpenForm formName
    OpenFormByName = True
End Function

Public Function OpenFormFiltered(ByVal formName As String, ByVal whereClause As String) As Boolean
    SaveActiveRecord
    DoCmd.OpenForm formName, , , whereClause
    OpenFormFiltered = True
End Function

' Opens a form filtered to keyField = keyValue and makes keyValue the default
' for new records, so records added there link back automatically.
Public Function OpenFormWhere(ByVal formName As String, ByVal keyField As String, ByVal keyValue As Variant) As Boolean
    If IsNull(keyValue) Then
        MsgBox "Save the current record first.", vbInformation, "EnviroLIMS"
        Exit Function
    End If
    SaveActiveRecord
    DoCmd.OpenForm formName, , , keyField & "=" & keyValue
    On Error Resume Next
    Forms(formName).Controls(keyField).DefaultValue = CStr(keyValue)
    OpenFormWhere = True
End Function

Public Function OpenQueryByName(ByVal queryName As String) As Boolean
    SaveActiveRecord
    DoCmd.OpenQuery queryName, acViewNormal, acReadOnly
    OpenQueryByName = True
End Function

Public Function OpenReportByName(ByVal reportName As String) As Boolean
    SaveActiveRecord
    DoCmd.OpenReport reportName, acViewPreview
    OpenReportByName = True
End Function

Public Function OpenReportWhere(ByVal reportName As String, ByVal keyField As String, ByVal keyValue As Variant) As Boolean
    If IsNull(keyValue) Then
        MsgBox "Save the current record first.", vbInformation, "EnviroLIMS"
        Exit Function
    End If
    SaveActiveRecord
    DoCmd.OpenReport reportName, acViewPreview, , keyField & "=" & keyValue
    OpenReportWhere = True
End Function

' Prompts for a work order number and previews its Certificate of Analysis.
Public Function PrintCOAPrompt() As Boolean
    Dim woNum As String, woID As Variant
    woNum = Trim$(InputBox("Enter the work order number for the Certificate of Analysis:", "EnviroLIMS"))
    If Len(woNum) = 0 Then Exit Function
    woID = DLookup("WorkOrderID", "tblWorkOrders", "WorkOrderNumber=" & SqlText(woNum))
    If IsNull(woID) Then
        MsgBox "Work order " & woNum & " was not found.", vbExclamation, "EnviroLIMS"
        Exit Function
    End If
    DoCmd.OpenReport "rptCertificateOfAnalysis", acViewPreview, , "WorkOrderID=" & woID
    PrintCOAPrompt = True
End Function

Public Function RefreshDashboard() As Boolean
    On Error Resume Next
    Screen.ActiveForm.Recalc
    RefreshDashboard = True
End Function

' ----------------------------------------------------------------------------
'  Numbering
' ----------------------------------------------------------------------------

' Next work order number in the form WO2611-001 (prefix + yymm + sequence).
Public Function NextWorkOrderNumber() As String
    Dim prefix As String, rs As DAO.Recordset, n As Long, maxN As Long
    prefix = LimsSetting("WorkOrderPrefix", "WO") & Format$(Date, "yymm") & "-"
    Set rs = CurrentDb.OpenRecordset("SELECT WorkOrderNumber FROM tblWorkOrders WHERE WorkOrderNumber LIKE " & SqlText(prefix & "*"), dbOpenSnapshot)
    Do While Not rs.EOF
        n = Val(Mid$(rs!WorkOrderNumber, Len(prefix) + 1))
        If n > maxN Then maxN = n
        rs.MoveNext
    Loop
    rs.Close
    NextWorkOrderNumber = prefix & Format$(maxN + 1, "000")
End Function

' Gives every sample on the work order without a Lab Sample ID the next
' number in sequence: <WorkOrderNumber>-01, -02, ...
Public Function AssignLabSampleIDs(ByVal workOrderID As Variant, Optional ByVal quiet As Boolean = False) As Long
    Dim db As DAO.Database, rs As DAO.Recordset, woNum As String, maxN As Long, n As Long, cnt As Long
    If IsNull(workOrderID) Then Exit Function
    SaveActiveRecord
    Set db = CurrentDb
    woNum = Nz(DLookup("WorkOrderNumber", "tblWorkOrders", "WorkOrderID=" & workOrderID), "")
    If Len(woNum) = 0 Then Exit Function

    Set rs = db.OpenRecordset("SELECT LabSampleID FROM tblSamples WHERE WorkOrderID=" & workOrderID & " AND LabSampleID LIKE " & SqlText(woNum & "-*"), dbOpenSnapshot)
    Do While Not rs.EOF
        n = Val(Mid$(rs!LabSampleID, Len(woNum) + 2))
        If n > maxN Then maxN = n
        rs.MoveNext
    Loop
    rs.Close

    Set rs = db.OpenRecordset("SELECT SampleID, LabSampleID FROM tblSamples WHERE WorkOrderID=" & workOrderID & " AND LabSampleID Is Null ORDER BY SampleID", dbOpenDynaset)
    Do While Not rs.EOF
        maxN = maxN + 1
        rs.Edit
        rs!LabSampleID = woNum & "-" & Format$(maxN, "00")
        rs.Update
        cnt = cnt + 1
        rs.MoveNext
    Loop
    rs.Close

    If cnt > 0 Then LogAudit "tblWorkOrders", workOrderID, "Assign Lab IDs", cnt & " sample(s)"
    RequeryActive
    If Not quiet Then MsgBox cnt & " lab sample ID(s) assigned.", vbInformation, "EnviroLIMS"
    AssignLabSampleIDs = cnt
End Function

' ----------------------------------------------------------------------------
'  Test / result generation
' ----------------------------------------------------------------------------

' Creates placeholder result rows (one per analyte in the method) for every
' test matching the WHERE clause, copies the method MDL/RL/units, and sets
' each test's due date from the method turnaround time.
Private Function GenerateResultRows(ByVal sampleWhere As String) As Long
    Dim db As DAO.Database, sql As String
    Set db = CurrentDb

    sql = "INSERT INTO tblResults (TestID, AnalyteID, Units, MDL, RL, DilutionFactor, Reviewed) " & _
          "SELECT t.TestID, ma.AnalyteID, ma.Units, ma.MDL, ma.RL, 1, False " & _
          "FROM (tblSampleTests AS t INNER JOIN tblSamples AS s ON t.SampleID = s.SampleID) " & _
          "INNER JOIN tblMethodAnalytes AS ma ON t.MethodID = ma.MethodID " & _
          "WHERE " & sampleWhere & " AND NOT EXISTS (SELECT r.ResultID FROM tblResults AS r " & _
          "WHERE r.TestID = t.TestID AND r.AnalyteID = ma.AnalyteID)"
    db.Execute sql, dbFailOnError
    GenerateResultRows = db.RecordsAffected

    sql = "UPDATE ((tblSampleTests AS t INNER JOIN tblSamples AS s ON t.SampleID = s.SampleID) " & _
          "INNER JOIN tblWorkOrders AS w ON s.WorkOrderID = w.WorkOrderID) " & _
          "INNER JOIN tblMethods AS mt ON t.MethodID = mt.MethodID " & _
          "SET t.DueDate = DateAdd('d', Nz(mt.TurnaroundDays, 10), Nz(w.ReceivedDate, Now())) " & _
          "WHERE t.DueDate Is Null AND " & sampleWhere
    db.Execute sql, dbFailOnError
End Function

Public Function CreateResultRows(ByVal sampleID As Variant, Optional ByVal quiet As Boolean = False) As Long
    Dim n As Long
    If IsNull(sampleID) Then Exit Function
    SaveActiveRecord
    n = GenerateResultRows("s.SampleID=" & sampleID)
    RequeryActive
    If Not quiet Then MsgBox n & " result row(s) created.", vbInformation, "EnviroLIMS"
    CreateResultRows = n
End Function

Public Function CreateResultRowsForWorkOrder(ByVal workOrderID As Variant, Optional ByVal quiet As Boolean = False) As Long
    Dim n As Long
    If IsNull(workOrderID) Then Exit Function
    SaveActiveRecord
    n = GenerateResultRows("s.WorkOrderID=" & workOrderID)
    RequeryActive
    If Not quiet Then MsgBox n & " result row(s) created.", vbInformation, "EnviroLIMS"
    CreateResultRowsForWorkOrder = n
End Function

' Adds a method to every sample on the work order whose matrix class
' (Aqueous / Solid / Air) matches the method, then creates result rows.
Public Function AddMethodToWorkOrder(ByVal workOrderID As Variant, ByVal methodID As Variant, Optional ByVal quiet As Boolean = False) As Long
    Dim db As DAO.Database, sql As String, cls As String, n As Long
    If IsNull(workOrderID) Then Exit Function
    If IsNull(methodID) Then
        MsgBox "Choose a method in the 'Add method' box first.", vbInformation, "EnviroLIMS"
        Exit Function
    End If
    SaveActiveRecord
    Set db = CurrentDb
    cls = Nz(DLookup("MatrixClass", "tblMethods", "MethodID=" & methodID), "")

    sql = "INSERT INTO tblSampleTests (SampleID, MethodID, Status) " & _
          "SELECT s.SampleID, " & methodID & ", 'Pending' " & _
          "FROM tblSamples AS s INNER JOIN tblMatrices AS mx ON s.MatrixID = mx.MatrixID " & _
          "WHERE s.WorkOrderID=" & workOrderID & _
          " AND NOT EXISTS (SELECT t.TestID FROM tblSampleTests AS t WHERE t.SampleID = s.SampleID AND t.MethodID=" & methodID & ")"
    If Len(cls) > 0 Then sql = sql & " AND mx.MatrixClass=" & SqlText(cls)
    db.Execute sql, dbFailOnError
    n = db.RecordsAffected

    GenerateResultRows "s.WorkOrderID=" & workOrderID
    UpdateWorkOrderStatus workOrderID
    LogAudit "tblWorkOrders", workOrderID, "Add Method", "MethodID " & methodID & " added to " & n & " sample(s)"
    RequeryActive
    If Not quiet Then
        If n = 0 And Len(cls) > 0 Then
            MsgBox "No tests added. Either every sample already has this method, or no sample has a " & cls & " matrix.", vbInformation, "EnviroLIMS"
        Else
            MsgBox "Method added to " & n & " sample(s).", vbInformation, "EnviroLIMS"
        End If
    End If
    AddMethodToWorkOrder = n
End Function

' ----------------------------------------------------------------------------
'  Result entry
' ----------------------------------------------------------------------------

' U = below MDL (not detected), J = between MDL and RL (estimated).
Public Function AutoQualifier(ByVal v As Variant, ByVal mdl As Variant, ByVal rl As Variant) As String
    If IsNull(v) Then Exit Function
    If Not IsNull(mdl) Then
        If v < mdl Then
            AutoQualifier = "U"
            Exit Function
        End If
    End If
    If Not IsNull(rl) Then
        If v < rl Then AutoQualifier = "J"
    End If
End Function

' AfterUpdate handler for ResultValue on the results subform.
Public Function ApplyQualifier() As Boolean
    Dim f As Form, p As Form, cur As String
    On Error GoTo EH
    Set f = Screen.ActiveControl.Parent

    cur = Nz(f!Qualifier, "")
    ' Only replace automatic flags; keep anything the analyst typed (B, H, E...)
    If cur = "" Or cur = "U" Or cur = "J" Then
        cur = AutoQualifier(f!ResultValue, f!MDL, f!RL)
        If Len(cur) = 0 Then f!Qualifier = Null Else f!Qualifier = cur
    End If
    If IsNull(f!AnalysisDate) And Not IsNull(f!ResultValue) Then f!AnalysisDate = Now()

    ' Move the parent test to In Progress
    On Error Resume Next
    Set p = f.Parent
    If Not p Is Nothing Then
        If Nz(p!Status, "") = "Pending" Then p!Status = "In Progress"
        If IsNull(p!AnalysisDate) Then p!AnalysisDate = Now()
    End If
    ApplyQualifier = True
    Exit Function
EH:
    ApplyQualifier = False
End Function

' Marks all results for a test as reviewed and the test as Reviewed.
Public Function ReviewTest(ByVal testID As Variant, Optional ByVal quiet As Boolean = False) As Boolean
    Dim db As DAO.Database, blanks As Long, woID As Variant
    If IsNull(testID) Then Exit Function
    SaveActiveRecord
    Set db = CurrentDb

    blanks = DCount("*", "tblResults", "TestID=" & testID & " AND ResultValue Is Null")
    If blanks > 0 And Not quiet Then
        If MsgBox(blanks & " result(s) on this test are still blank. Mark the test reviewed anyway?", _
                  vbYesNo + vbQuestion, "EnviroLIMS") = vbNo Then Exit Function
    End If

    db.Execute "UPDATE tblResults SET Reviewed=True, ReviewedBy=" & SqlText(CurrentLimsUser()) & _
               ", ReviewedDate=Now() WHERE TestID=" & testID, dbFailOnError
    db.Execute "UPDATE tblSampleTests SET Status='Reviewed', AnalysisDate=IIf(AnalysisDate Is Null, Now(), AnalysisDate) " & _
               "WHERE TestID=" & testID, dbFailOnError
    LogAudit "tblSampleTests", testID, "Reviewed"

    woID = DLookup("SampleID", "tblSampleTests", "TestID=" & testID)
    If Not IsNull(woID) Then woID = DLookup("WorkOrderID", "tblSamples", "SampleID=" & woID)
    If Not IsNull(woID) Then UpdateWorkOrderStatus woID

    RequeryActive
    If Not quiet Then MsgBox "Test marked as reviewed.", vbInformation, "EnviroLIMS"
    ReviewTest = True
End Function

' Rolls the test statuses up to the work order.
Public Sub UpdateWorkOrderStatus(ByVal workOrderID As Variant)
    Dim rs As DAO.Recordset, total As Long, reviewed As Long, pending As Long, cur As String, newStatus As String
    If IsNull(workOrderID) Then Exit Sub
    cur = Nz(DLookup("Status", "tblWorkOrders", "WorkOrderID=" & workOrderID), "")
    If cur = "Reported" Or cur = "Cancelled" Or cur = "On Hold" Then Exit Sub

    Set rs = CurrentDb.OpenRecordset( _
        "SELECT Count(t.TestID) AS Total, Sum(IIf(t.Status='Reviewed',1,0)) AS Rev, Sum(IIf(t.Status='Pending',1,0)) AS Pend " & _
        "FROM tblSampleTests AS t INNER JOIN tblSamples AS s ON t.SampleID = s.SampleID " & _
        "WHERE s.WorkOrderID=" & workOrderID, dbOpenSnapshot)
    total = Nz(rs!total, 0)
    reviewed = Nz(rs!Rev, 0)
    pending = Nz(rs!Pend, 0)
    rs.Close
    If total = 0 Then Exit Sub

    If reviewed = total Then
        newStatus = "Reviewed"
    ElseIf pending = total Then
        newStatus = "Received"
    Else
        newStatus = "In Progress"
    End If
    If newStatus <> cur Then
        CurrentDb.Execute "UPDATE tblWorkOrders SET Status=" & SqlText(newStatus) & " WHERE WorkOrderID=" & workOrderID, dbFailOnError
    End If
End Sub

' ----------------------------------------------------------------------------
'  Quality control
' ----------------------------------------------------------------------------

Public Function CalcRecovery(ByVal measured As Variant, ByVal spike As Variant, ByVal background As Variant) As Variant
    CalcRecovery = Null
    If IsNull(measured) Or IsNull(spike) Then Exit Function
    If spike = 0 Then Exit Function
    CalcRecovery = Round((measured - Nz(background, 0)) / spike * 100, 1)
End Function

Public Function CalcRPD(ByVal a As Variant, ByVal b As Variant) As Variant
    CalcRPD = Null
    If IsNull(a) Or IsNull(b) Then Exit Function
    If (a + b) = 0 Then
        CalcRPD = 0
    Else
        CalcRPD = Round(Abs(a - b) / ((a + b) / 2) * 100, 1)
    End If
End Function

' PASS / FAIL / INCOMPLETE / N/A for one QC result row.
'   Method Blank : measured must be <= UpperLimit (normally the RL)
'   Duplicate    : RPD(measured, duplicate) <= RPDLimit
'   LCS/MS/CCV...: %Recovery within LowerLimit..UpperLimit, and RPD <= RPDLimit
'                  when a duplicate value is supplied (LCSD / MSD)
Public Function EvalQC(ByVal qcType As Variant, ByVal measured As Variant, ByVal spike As Variant, ByVal background As Variant, _
                       ByVal dup As Variant, ByVal lowLimit As Variant, ByVal highLimit As Variant, ByVal rpdLimit As Variant) As String
    Dim rec As Variant, r As Variant, ok As Boolean
    Select Case UCase$(Nz(qcType, ""))
        Case "METHOD BLANK", "BLANK"
            If IsNull(measured) Then
                EvalQC = "INCOMPLETE"
            ElseIf IsNull(highLimit) Then
                EvalQC = "N/A"
            ElseIf measured <= highLimit Then
                EvalQC = "PASS"
            Else
                EvalQC = "FAIL"
            End If

        Case "DUPLICATE"
            r = CalcRPD(measured, dup)
            If IsNull(r) Then
                EvalQC = "INCOMPLETE"
            ElseIf IsNull(rpdLimit) Then
                EvalQC = "N/A"
            ElseIf r <= rpdLimit Then
                EvalQC = "PASS"
            Else
                EvalQC = "FAIL"
            End If

        Case ""
            EvalQC = ""

        Case Else
            rec = CalcRecovery(measured, spike, background)
            If IsNull(rec) Then
                EvalQC = "INCOMPLETE"
                Exit Function
            End If
            If IsNull(lowLimit) And IsNull(highLimit) Then
                EvalQC = "N/A"
                Exit Function
            End If
            ok = True
            If Not IsNull(lowLimit) Then
                If rec < lowLimit Then ok = False
            End If
            If Not IsNull(highLimit) Then
                If rec > highLimit Then ok = False
            End If
            If Not IsNull(dup) And Not IsNull(rpdLimit) Then
                r = CalcRPD(measured, dup)
                If Not IsNull(r) Then
                    If r > rpdLimit Then ok = False
                End If
            End If
            EvalQC = IIf(ok, "PASS", "FAIL")
    End Select
End Function

' Adds one QC row per method analyte to a batch, pre-filled with the
' acceptance limits stored on the method.
Public Function AddQCRows(ByVal batchID As Variant, ByVal qcType As Variant, Optional ByVal quiet As Boolean = False) As Long
    Dim db As DAO.Database, methodID As Variant, lo As String, hi As String, rpd As String, sql As String
    If IsNull(batchID) Then Exit Function
    If IsNull(qcType) Then
        MsgBox "Choose a QC type first.", vbInformation, "EnviroLIMS"
        Exit Function
    End If
    SaveActiveRecord
    Set db = CurrentDb
    methodID = DLookup("MethodID", "tblQCBatches", "BatchID=" & batchID)
    If IsNull(methodID) Then Exit Function

    Select Case UCase$(qcType)
        Case "METHOD BLANK"
            lo = "Null": hi = "ma.RL": rpd = "Null"
        Case "DUPLICATE"
            lo = "Null": hi = "Null": rpd = "ma.RPDLimit"
        Case "LCSD", "MSD"
            lo = "ma.RecoveryLow": hi = "ma.RecoveryHigh": rpd = "ma.RPDLimit"
        Case Else
            lo = "ma.RecoveryLow": hi = "ma.RecoveryHigh": rpd = "Null"
    End Select

    sql = "INSERT INTO tblQCResults (BatchID, QCType, AnalyteID, Units, LowerLimit, UpperLimit, RPDLimit) " & _
          "SELECT " & batchID & ", " & SqlText(qcType) & ", ma.AnalyteID, ma.Units, " & lo & ", " & hi & ", " & rpd & " " & _
          "FROM tblMethodAnalytes AS ma WHERE ma.MethodID=" & methodID
    db.Execute sql, dbFailOnError
    AddQCRows = db.RecordsAffected
    RequeryActive
    If Not quiet Then MsgBox AddQCRows & " " & qcType & " row(s) added.", vbInformation, "EnviroLIMS"
End Function
