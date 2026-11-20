Attribute VB_Name = "modBuildUI"
Option Compare Database
Option Explicit

' =============================================================================
'  EnviroLIMS - form and report generator
'
'  Forms are generated in code so the whole UI can be rebuilt (RebuildUI)
'  after schema changes. Field specs use the format
'       "FieldName|Caption|Type|Extra|DefaultValue"
'  Type:  T text   D date   M memo   K checkbox
'         C combo  (Extra = 2-column row source SQL: ID, display text)
'         V value list combo (Extra = "a;b;c")
'         Q 1-column query combo that also accepts typed values (Extra = SQL)
'         R read-only calculated text box (Extra = expression, no leading =)
'  Append L to the type to lock the control (e.g. "CL").
' =============================================================================

Private Const ROWH As Long = 315
Private Const ROWGAP As Long = 400
Private Const RPT_W As Long = 10800       ' 7.5 inches

' --- Combo box row sources ---
Private Const SQL_CLIENTS As String = "SELECT ClientID, ClientName FROM tblClients ORDER BY ClientName"
Private Const SQL_PROJECTS As String = "SELECT p.ProjectID, c.ClientName & ' - ' & p.ProjectName FROM tblProjects AS p " & _
    "INNER JOIN tblClients AS c ON p.ClientID = c.ClientID ORDER BY c.ClientName, p.ProjectName"
Private Const SQL_LOCATIONS As String = "SELECT l.LocationID, l.LocationCode & ' (' & p.ProjectName & ')' FROM tblLocations AS l " & _
    "INNER JOIN tblProjects AS p ON l.ProjectID = p.ProjectID ORDER BY l.LocationCode"
Private Const SQL_MATRICES As String = "SELECT MatrixID, MatrixName FROM tblMatrices ORDER BY MatrixName"
Private Const SQL_ANALYTES As String = "SELECT AnalyteID, AnalyteName FROM tblAnalytes ORDER BY AnalyteName"
Private Const SQL_METHODS As String = "SELECT MethodID, MethodCode & ' - ' & Nz(MethodName,'') FROM tblMethods ORDER BY MethodCode"
Private Const SQL_ACTIVE_METHODS As String = "SELECT MethodID, MethodCode & ' - ' & Nz(MethodName,'') FROM tblMethods WHERE Active = True ORDER BY MethodCode"
Private Const SQL_ANALYSTS As String = "SELECT AnalystID, Initials & ' - ' & Nz(FirstName,'') & ' ' & Nz(LastName,'') FROM tblAnalysts ORDER BY Initials"
Private Const SQL_INSTRUMENTS As String = "SELECT InstrumentID, InstrumentName FROM tblInstruments ORDER BY InstrumentName"
Private Const SQL_BATCHES As String = "SELECT b.BatchID, b.BatchNumber & ' (' & m.MethodCode & ')' FROM tblQCBatches AS b " & _
    "INNER JOIN tblMethods AS m ON b.MethodID = m.MethodID ORDER BY b.BatchNumber DESC"
Private Const SQL_WORKORDERS As String = "SELECT w.WorkOrderID, w.WorkOrderNumber & ' - ' & c.ClientName FROM tblWorkOrders AS w " & _
    "INNER JOIN tblClients AS c ON w.ClientID = c.ClientID ORDER BY w.WorkOrderNumber DESC"
Private Const SQL_SAMPLES As String = "SELECT SampleID, Nz(LabSampleID,'(unassigned)') & ' / ' & Nz(ClientSampleID,'') FROM tblSamples ORDER BY LabSampleID"
Private Const SQL_LIMITSETS As String = "SELECT DISTINCT LimitSet FROM tblRegulatoryLimits ORDER BY LimitSet"

' --- Value lists ---
Private Const VL_WO_STATUS As String = "Received;In Progress;Reviewed;Reported;On Hold;Cancelled"
Private Const VL_SAMPLE_STATUS As String = "Received;In Progress;Complete;Disposed;On Hold;Cancelled"
Private Const VL_TEST_STATUS As String = "Pending;In Progress;Complete;Reviewed;Cancelled"
Private Const VL_SAMPLE_TYPE As String = "Grab;Composite;Field Blank;Trip Blank;Equipment Blank;Field Duplicate"
Private Const VL_QC_TYPE As String = "Method Blank;LCS;LCSD;MS;MSD;Duplicate;ICV;CCV"
Private Const VL_QUALIFIER As String = "U;J;B;H;E;D;*"
Private Const VL_BATCH_STATUS As String = "Open;Complete;Reviewed"
Private Const VL_CUSTODY As String = "Relinquished by Sampler;Received by Lab;Transferred to Storage;Removed for Analysis;Returned to Storage;Disposed"
Private Const VL_LOCTYPE As String = "Entry Point;Distribution Tap;Monitoring Well;Surface Water;Outfall;Influent;Effluent;Soil Boring;Sediment"
Private Const VL_MATRIXCLASS As String = "Aqueous;Solid;Air"

Private Const UI_FORMS As String = _
    "frmMainMenu,frmWorkOrders,sfrmWOSamples,frmSamples,sfrmSampleTests,frmResultEntry,sfrmResults," & _
    "frmQCBatches,sfrmQCResults,frmClients,sfrmProjects,frmProjects,sfrmLocations,frmMethods,sfrmMethodAnalytes," & _
    "frmAnalytes,frmMatrices,frmAnalysts,frmInstruments,frmRegulatoryLimits,frmSettings,frmCustodyLog,frmAuditLog"
Private Const UI_REPORTS As String = _
    "rptCertificateOfAnalysis,rptChainOfCustody,rptExceedances,rptHoldTimes,rptQCSummary,rptWorkOrderStatus"

' =============================================================================
'  Entry point
' =============================================================================

Public Sub BuildAllUI()
    On Error GoTo EH
    Application.Echo False
    DeleteUI
    BuildSetupForms
    BuildOperationsForms
    BuildReports
    BuildMainMenu
    Application.Echo True
    Exit Sub
EH:
    Dim n As Long, d As String
    n = Err.Number: d = Err.Description
    Application.Echo True
    Err.Raise n, "BuildAllUI", d
End Sub

Private Sub DeleteUI()
    Dim i As Long, names() As String
    names = Split(UI_FORMS, ",")
    For i = LBound(names) To UBound(names)
        DeleteObj acForm, names(i)
    Next i
    names = Split(UI_REPORTS, ",")
    For i = LBound(names) To UBound(names)
        DeleteObj acReport, names(i)
    Next i
End Sub

' =============================================================================
'  Generic helpers
' =============================================================================

Private Sub DeleteObj(ByVal objType As AcObjectType, ByVal objName As String)
    On Error Resume Next
    DoCmd.Close objType, objName, acSaveNo
    DoCmd.DeleteObject objType, objName
    Err.Clear
End Sub

Private Sub FinishObject(ByVal objType As AcObjectType, ByVal tmpName As String, ByVal finalName As String)
    DoCmd.Save objType, tmpName
    DoCmd.Close objType, tmpName, acSaveYes
    DeleteObj objType, finalName
    DoCmd.Rename finalName, objType, tmpName
End Sub

Private Function NewForm(ByVal src As String, ByVal cap As String, ByVal view As Integer) As Form
    Dim f As Form
    Set f = CreateForm()
    f.RecordSource = src
    f.Caption = cap
    f.DefaultView = view
    f.AutoCenter = True
    f.DividingLines = False
    Set NewForm = f
End Function

Private Function AddLabel(ByVal frm As Form, ByVal text As String, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, _
                          Optional ByVal fontSize As Integer = 10, Optional ByVal bold As Boolean = False) As Control
    Dim c As Control
    Set c = CreateControl(frm.Name, acLabel, acDetail, "", "", x, y, w, h)
    c.Caption = text
    c.FontSize = fontSize
    c.FontBold = bold
    Set AddLabel = c
End Function

Private Function AddButton(ByVal frm As Form, ByVal cap As String, ByVal onClickExpr As String, ByVal x As Long, ByVal y As Long, _
                           Optional ByVal w As Long = 2200, Optional ByVal h As Long = 420) As Control
    Dim c As Control
    Set c = CreateControl(frm.Name, acCommandButton, acDetail, "", "", x, y, w, h)
    c.Caption = cap
    c.OnClick = onClickExpr
    Set AddButton = c
End Function

' Adds a row of buttons from specs "Caption|=Expression"; returns the next free Y.
Private Function AddButtonRow(ByVal frm As Form, buttons As Variant, ByVal x0 As Long, ByVal y As Long) As Long
    Dim i As Long, p() As String, x As Long
    x = x0
    For i = LBound(buttons) To UBound(buttons)
        p = Split(CStr(buttons(i)), "|")
        AddButton frm, p(0), p(1), x, y
        x = x + 2350
        If x > 9000 And i < UBound(buttons) Then
            x = x0
            y = y + 520
        End If
    Next i
    AddButtonRow = y + 560
End Function

Private Function AddField(ByVal frm As Form, ByVal spec As String, ByVal x As Long, ByVal y As Long, _
                          ByVal lblW As Long, ByVal ctlW As Long) As Control
    Dim p() As String, fld As String, cap As String, typ As String, kind As String, extra As String
    Dim ctl As Control, lbl As Control, isDatasheet As Boolean

    p = Split(spec, "|")
    fld = p(0): cap = p(1): typ = UCase$(p(2))
    If UBound(p) >= 3 Then extra = p(3)
    kind = Left$(typ, 1)
    isDatasheet = (frm.DefaultView = 2)

    Select Case kind
        Case "C", "V", "Q"
            Set ctl = CreateControl(frm.Name, acComboBox, acDetail, "", fld, x + lblW, y, ctlW, ROWH)
            If kind = "C" Then
                ctl.RowSourceType = "Table/Query"
                ctl.RowSource = extra
                ctl.ColumnCount = 2
                ctl.BoundColumn = 1
                ctl.ColumnWidths = "0;" & IIf(ctlW < 2880, 2880, ctlW)
                ctl.ListWidth = IIf(ctlW < 2880, 2880, ctlW) + 250
                ctl.LimitToList = True
            ElseIf kind = "Q" Then
                ctl.RowSourceType = "Table/Query"
                ctl.RowSource = extra
                ctl.ColumnCount = 1
                ctl.LimitToList = False
            Else
                ctl.RowSourceType = "Value List"
                ctl.RowSource = extra
                ctl.LimitToList = False
            End If
        Case "K"
            Set ctl = CreateControl(frm.Name, acCheckBox, acDetail, "", fld, x + lblW, y + 40, 260, 240)
        Case "R"
            Set ctl = CreateControl(frm.Name, acTextBox, acDetail, "", "", x + lblW, y, ctlW, ROWH)
            ctl.ControlSource = "=" & extra
            ctl.Locked = True
            ctl.TabStop = False
            ctl.BackColor = RGB(242, 242, 242)
        Case "M"
            Set ctl = CreateControl(frm.Name, acTextBox, acDetail, "", fld, x + lblW, y, ctlW, 900)
            ctl.EnterKeyBehavior = True
            ctl.ScrollBars = 2
        Case Else
            Set ctl = CreateControl(frm.Name, acTextBox, acDetail, "", fld, x + lblW, y, ctlW, ROWH)
    End Select
    ctl.Name = fld

    If UBound(p) >= 4 Then
        If Len(p(4)) > 0 Then ctl.DefaultValue = p(4)
    End If
    If kind <> "R" And InStr(2, typ, "L") > 0 Then
        ctl.Locked = True
        ctl.TabStop = False
        ctl.BackColor = RGB(242, 242, 242)
    End If

    If isDatasheet Then
        On Error Resume Next
        Select Case kind
            Case "C": ctl.ColumnWidth = 2600
            Case "M": ctl.ColumnWidth = 3200
            Case "K": ctl.ColumnWidth = 900
            Case Else: ctl.ColumnWidth = 1500
        End Select
        On Error GoTo 0
    End If

    Set lbl = CreateControl(frm.Name, acLabel, acDetail, ctl.Name, "", x, y, lblW - 80, ROWH)
    lbl.Caption = cap
    Set AddField = ctl
End Function

' Lays fields out in columns; returns the Y below the lowest field.
Private Function LayoutFields(ByVal frm As Form, specs As Variant, ByVal cols As Integer, ByVal x0 As Long, ByVal y0 As Long, _
                              ByVal lblW As Long, ByVal ctlW As Long) As Long
    Dim n As Long, perCol As Long, i As Long, c As Integer, y As Long, maxY As Long, k As Long, typ As String
    n = UBound(specs) - LBound(specs) + 1
    perCol = -Int(-n / cols)
    k = LBound(specs)
    maxY = y0
    For c = 0 To cols - 1
        y = y0
        For i = 1 To perCol
            If k > UBound(specs) Then Exit For
            AddField frm, CStr(specs(k)), x0 + c * (lblW + ctlW + 300), y, lblW, ctlW
            typ = UCase$(Split(CStr(specs(k)), "|")(2))
            If Left$(typ, 1) = "M" Then y = y + 1000 Else y = y + ROWGAP
            k = k + 1
        Next i
        If y > maxY Then maxY = y
    Next c
    LayoutFields = maxY
End Function

Private Function AddSubform(ByVal frm As Form, ByVal sourceForm As String, ByVal master As String, ByVal child As String, _
                            ByVal cap As String, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long) As Control
    Dim c As Control
    AddLabel frm, cap, x, y, w, 320, 11, True
    Set c = CreateControl(frm.Name, acSubform, acDetail, "", "", x, y + 360, w, h)
    c.Name = "sub" & sourceForm
    c.SourceObject = "Form." & sourceForm
    c.LinkMasterFields = master
    c.LinkChildFields = child
    Set AddSubform = c
End Function

Private Sub RedWhen(ByVal ctl As Control, ByVal valueText As String)
    Dim fc As Object
    On Error Resume Next
    Set fc = ctl.FormatConditions.Add(acFieldValue, acEqual, """" & valueText & """")
    fc.ForeColor = RGB(192, 0, 0)
    fc.FontBold = True
End Sub

' Builds a datasheet form (left open in design view; call FinishObject).
Private Function StartDatasheet(ByVal src As String, ByVal cap As String, specs As Variant, _
                                Optional ByVal allowAdd As Boolean = True, Optional ByVal allowEdit As Boolean = True) As Form
    Dim frm As Form, i As Long, y As Long
    Set frm = NewForm(src, cap, 2)
    y = 100
    For i = LBound(specs) To UBound(specs)
        AddField frm, CStr(specs(i)), 100, y, 2000, 2600
        y = y + 1000
    Next i
    frm.AllowAdditions = allowAdd
    frm.AllowDeletions = allowAdd
    frm.AllowEdits = allowEdit
    Set StartDatasheet = frm
End Function

Private Sub MakeDatasheet(ByVal formName As String, ByVal src As String, ByVal cap As String, specs As Variant, _
                          Optional ByVal allowAdd As Boolean = True, Optional ByVal allowEdit As Boolean = True)
    Dim frm As Form
    Set frm = StartDatasheet(src, cap, specs, allowAdd, allowEdit)
    FinishObject acForm, frm.Name, formName
End Sub

' Builds a single-record form with a title and a 2-column field grid.
' Returns the form; the next free Y position is returned in nextY.
Private Function StartSingle(ByVal src As String, ByVal cap As String, specs As Variant, ByRef nextY As Long) As Form
    Dim frm As Form
    Set frm = NewForm(src, cap, 0)
    frm.Width = 11000
    AddLabel frm, cap, 200, 120, 8000, 480, 16, True
    nextY = LayoutFields(frm, specs, 2, 200, 750, 1700, 3000) + 100
    Set StartSingle = frm
End Function

' =============================================================================
'  Setup / reference forms
' =============================================================================

Private Sub BuildSetupForms()
    Dim frm As Form, y As Long

    MakeDatasheet "frmSettings", "tblSettings", "Settings", _
        Array("SettingName|Setting|TL", "SettingValue|Value|T", "Description|Description|TL"), False

    MakeDatasheet "frmMatrices", "SELECT * FROM tblMatrices ORDER BY MatrixName", "Matrices", _
        Array("MatrixCode|Code|T", "MatrixName|Matrix|T", "MatrixClass|Class|V|" & VL_MATRIXCLASS)

    MakeDatasheet "frmAnalytes", "SELECT * FROM tblAnalytes ORDER BY AnalyteGroup, AnalyteName", "Analytes", _
        Array("AnalyteName|Analyte|T", "CASNumber|CAS Number|T", "AnalyteGroup|Group|T")

    MakeDatasheet "frmAnalysts", "SELECT * FROM tblAnalysts ORDER BY Initials", "Analysts", _
        Array("Initials|Initials|T", "FirstName|First Name|T", "LastName|Last Name|T", "Title|Title|T", "Active|Active|K")

    MakeDatasheet "frmInstruments", "SELECT * FROM tblInstruments ORDER BY InstrumentName", "Instruments", _
        Array("InstrumentName|Instrument|T", "Model|Model|T", "SerialNumber|Serial #|T", _
              "LastCalibration|Last Calibration|D", "NextCalibrationDue|Calibration Due|D", "Active|Active|K")

    MakeDatasheet "frmRegulatoryLimits", "SELECT * FROM tblRegulatoryLimits ORDER BY LimitSet, AnalyteID", "Regulatory Limits", _
        Array("LimitSet|Limit Set|T", "AnalyteID|Analyte|C|" & SQL_ANALYTES, "MatrixID|Matrix|C|" & SQL_MATRICES, _
              "LimitValue|Limit|T", "Units|Units|T", "LimitType|Type|T", "Source|Source|T")

    MakeDatasheet "frmAuditLog", "SELECT * FROM tblAuditLog ORDER BY EventTime DESC", "Audit Log", _
        Array("EventTime|Time|D", "UserName|User|T", "TableName|Table|T", "RecordID|Record ID|T", _
              "ActionTaken|Action|T", "Details|Details|T"), False, False

    MakeDatasheet "frmCustodyLog", "SELECT * FROM tblCustodyLog ORDER BY EventDate", "Chain of Custody Log", _
        Array("WorkOrderID|Work Order|C|" & SQL_WORKORDERS, "EventDate|Date/Time|D", "CustodyAction|Action|V|" & VL_CUSTODY, _
              "RelinquishedBy|Relinquished By|T", "ReceivedBy|Received By|T", "StorageLocation|Location|T", "Comments|Comments|T")

    ' --- Clients with their projects ---
    Set frm = StartDatasheet("SELECT * FROM tblProjects ORDER BY ProjectName", "Projects", _
        Array("ProjectName|Project|T", "ProjectCode|Code|T", "SiteAddress|Site|T", "LimitSet|Limit Set|Q|" & SQL_LIMITSETS, _
              "PermitNumber|Permit #|T", "StartDate|Start Date|D", "Active|Active|K"))
    frm.Controls("ProjectName").OnDblClick = "=OpenFormWhere(""frmProjects"",""ProjectID"",[ProjectID])"
    FinishObject acForm, frm.Name, "sfrmProjects"

    Set frm = StartSingle("SELECT * FROM tblClients ORDER BY ClientName", "Clients", _
        Array("ClientName|Client Name|T", "ContactName|Contact|T", "Phone|Phone|T", "Email|Email|T", "Active|Active|K", _
              "Address|Address|T", "City|City|T", "State|State|T", "PostalCode|Postal Code|T", "Notes|Notes|M"), y)
    AddSubform frm, "sfrmProjects", "ClientID", "ClientID", "Projects (double-click a project to open it)", 200, y, 10400, 2600
    FinishObject acForm, frm.Name, "frmClients"

    ' --- Projects with their sampling locations ---
    MakeDatasheet "sfrmLocations", "SELECT * FROM tblLocations ORDER BY LocationCode", "Sampling Locations", _
        Array("LocationCode|Location Code|T", "Description|Description|T", "LocationType|Type|V|" & VL_LOCTYPE, _
              "Latitude|Latitude|T", "Longitude|Longitude|T", "Active|Active|K")

    Set frm = StartSingle("SELECT * FROM tblProjects ORDER BY ProjectName", "Projects", _
        Array("ClientID|Client|C|" & SQL_CLIENTS, "ProjectName|Project Name|T", "ProjectCode|Project Code|T", _
              "SiteAddress|Site Address|T", "Active|Active|K", "LimitSet|Limit Set|Q|" & SQL_LIMITSETS, _
              "PermitNumber|Permit #|T", "StartDate|Start Date|D", "Notes|Notes|M"), y)
    AddSubform frm, "sfrmLocations", "ProjectID", "ProjectID", "Sampling Locations", 200, y, 10400, 2600
    FinishObject acForm, frm.Name, "frmProjects"

    ' --- Methods with their analytes ---
    MakeDatasheet "sfrmMethodAnalytes", "SELECT * FROM tblMethodAnalytes ORDER BY MethodAnalyteID", "Method Analytes", _
        Array("AnalyteID|Analyte|C|" & SQL_ANALYTES, "Units|Units|T", "MDL|MDL|T", "RL|RL|T", _
              "RecoveryLow|%Rec Low|T", "RecoveryHigh|%Rec High|T", "RPDLimit|RPD Limit|T")

    Set frm = StartSingle("SELECT * FROM tblMethods ORDER BY MethodCode", "Methods", _
        Array("MethodCode|Method Code|T", "MethodName|Method Name|T", "MatrixClass|Matrix Class|V|" & VL_MATRIXCLASS, _
              "Active|Active|K", "HoldTimeHours|Hold Time (hours)|T", "TurnaroundDays|Turnaround (days)|T", _
              "Preservation|Preservation|T", "ContainerType|Container|T"), y)
    AddSubform frm, "sfrmMethodAnalytes", "MethodID", "MethodID", "Analytes, detection limits and QC acceptance criteria", 200, y, 10400, 3200
    FinishObject acForm, frm.Name, "frmMethods"
End Sub

' =============================================================================
'  Operations forms: work orders, samples, results, QC
' =============================================================================

Private Sub BuildOperationsForms()
    Dim frm As Form, y As Long, c As Control

    ' --- Samples subform on the work order (sample login grid) ---
    Set frm = StartDatasheet("SELECT * FROM tblSamples ORDER BY SampleID", "Samples", _
        Array("LabSampleID|Lab Sample ID|T", "ClientSampleID|Client Sample ID|T", "LocationID|Location|C|" & SQL_LOCATIONS, _
              "MatrixID|Matrix|C|" & SQL_MATRICES, "SampleType|Type|V|" & VL_SAMPLE_TYPE, "CollectedDate|Collected|D", _
              "CollectedBy|Collected By|T", "Containers|Containers|T", "Preservative|Preservative|T", _
              "Status|Status|V|" & VL_SAMPLE_STATUS, "calcTests|Tests|R|DCount(""*"",""tblSampleTests"",""SampleID="" & Nz([SampleID],0))"))
    frm.Controls("LabSampleID").OnDblClick = "=OpenFormWhere(""frmSamples"",""SampleID"",[SampleID])"
    frm.Controls("ClientSampleID").OnDblClick = "=OpenFormWhere(""frmSamples"",""SampleID"",[SampleID])"
    FinishObject acForm, frm.Name, "sfrmWOSamples"

    ' --- Work orders ---
    Set frm = StartSingle("SELECT * FROM tblWorkOrders ORDER BY WorkOrderNumber DESC", "Work Orders / Sample Receiving", _
        Array("WorkOrderNumber|Work Order #|T||NextWorkOrderNumber()", "ClientID|Client|C|" & SQL_CLIENTS, _
              "ProjectID|Project|C|" & SQL_PROJECTS, "ReceivedDate|Received|D", "ReceivedByID|Received By|C|" & SQL_ANALYSTS, _
              "DueDate|Due Date|D||Date()+10", "Status|Status|V|" & VL_WO_STATUS, _
              "PONumber|PO Number|T", "CoolerTempC|Cooler Temp (C)|T", "ReceivedOnIce|Received on Ice|K", _
              "CustodySealsIntact|Custody Seals Intact|K", "Notes|Notes|M"), y)
    y = AddButtonRow(frm, Array( _
        "Assign Lab IDs|=AssignLabSampleIDs([WorkOrderID])", _
        "Create Result Rows|=CreateResultRowsForWorkOrder([WorkOrderID])", _
        "Custody Log|=OpenFormWhere(""frmCustodyLog"",""WorkOrderID"",[WorkOrderID])", _
        "Print COA|=OpenReportWhere(""rptCertificateOfAnalysis"",""WorkOrderID"",[WorkOrderID])", _
        "Print Custody|=OpenReportWhere(""rptChainOfCustody"",""WorkOrderID"",[WorkOrderID])"), 200, y)
    AddLabel frm, "Add method to all matching samples:", 200, y, 3200, ROWH
    Set c = CreateControl(frm.Name, acComboBox, acDetail, "", "", 3450, y, 4200, ROWH)
    c.Name = "cboAddMethod"
    c.RowSourceType = "Table/Query"
    c.RowSource = SQL_ACTIVE_METHODS
    c.ColumnCount = 2
    c.BoundColumn = 1
    c.ColumnWidths = "0;4200"
    c.ListWidth = 4500
    c.LimitToList = True
    AddButton frm, "Add Method", "=AddMethodToWorkOrder([WorkOrderID],[cboAddMethod])", 7800, y - 50, 1800
    AddSubform frm, "sfrmWOSamples", "WorkOrderID", "WorkOrderID", _
        "Samples (enter samples, then Assign Lab IDs; double-click a sample to open it)", 200, y + 500, 10400, 3200
    FinishObject acForm, frm.Name, "frmWorkOrders"

    ' --- Tests subform on the sample ---
    Set frm = StartDatasheet("SELECT * FROM tblSampleTests ORDER BY TestID", "Tests", _
        Array("MethodID|Method|C|" & SQL_METHODS, "Status|Status|V|" & VL_TEST_STATUS, "AnalystID|Analyst|C|" & SQL_ANALYSTS, _
              "BatchID|QC Batch|C|" & SQL_BATCHES, "PrepDate|Prep Date|D", "AnalysisDate|Analysis Date|D", "DueDate|Due Date|D", _
              "calcHold|Hold Time|R|DLookup(""HoldStatus"",""qryHoldTimeStatus"",""TestID="" & Nz([TestID],0))"))
    frm.Controls("MethodID").OnDblClick = "=OpenFormWhere(""frmResultEntry"",""TestID"",[TestID])"
    RedWhen frm.Controls("calcHold"), "EXPIRED"
    RedWhen frm.Controls("calcHold"), "OUT OF HOLD"
    FinishObject acForm, frm.Name, "sfrmSampleTests"

    ' --- Samples ---
    Set frm = StartSingle("SELECT * FROM tblSamples ORDER BY LabSampleID", "Samples", _
        Array("WorkOrderID|Work Order|C|" & SQL_WORKORDERS, "LabSampleID|Lab Sample ID|T", "ClientSampleID|Client Sample ID|T", _
              "LocationID|Location|C|" & SQL_LOCATIONS, "MatrixID|Matrix|C|" & SQL_MATRICES, "SampleType|Sample Type|V|" & VL_SAMPLE_TYPE, _
              "CollectedDate|Collected|D", "CollectedBy|Collected By|T", "Containers|Containers|T", "Preservative|Preservative|T", _
              "FieldPH|Field pH|T", "FieldTempC|Field Temp (C)|T", "Status|Status|V|" & VL_SAMPLE_STATUS, _
              "DisposalDate|Disposal Date|D", "Notes|Notes|M"), y)
    y = AddButtonRow(frm, Array( _
        "Create Result Rows|=CreateResultRows([SampleID])", _
        "Enter Results|=OpenFormWhere(""frmResultEntry"",""SampleID"",[SampleID])", _
        "Open Work Order|=OpenFormWhere(""frmWorkOrders"",""WorkOrderID"",[WorkOrderID])"), 200, y)
    AddSubform frm, "sfrmSampleTests", "SampleID", "SampleID", _
        "Tests (add a method, then Create Result Rows; double-click a method to enter results)", 200, y, 10400, 2800
    FinishObject acForm, frm.Name, "frmSamples"

    ' --- Results subform ---
    Set frm = StartDatasheet("SELECT * FROM tblResults ORDER BY ResultID", "Results", _
        Array("AnalyteID|Analyte|CL|" & SQL_ANALYTES, "ResultValue|Result|T", "Qualifier|Qual|V|" & VL_QUALIFIER, _
              "Units|Units|T", "DilutionFactor|DF|T", "MDL|MDL|T", "RL|RL|T", "AnalystID|Analyst|C|" & SQL_ANALYSTS, _
              "AnalysisDate|Analyzed|D", _
              "calcLimit|Limit|R|DLookup(""LimitValue"",""qryResultsWithLimits"",""ResultID="" & Nz([ResultID],0))", _
              "calcFlag|Flag|R|DLookup(""ExceedFlag"",""qryResultsWithLimits"",""ResultID="" & Nz([ResultID],0))", _
              "Reviewed|Reviewed|KL", "Comments|Comments|T"), False, True)
    frm.Controls("ResultValue").AfterUpdate = "=ApplyQualifier()"
    RedWhen frm.Controls("calcFlag"), "EXCEEDS"
    FinishObject acForm, frm.Name, "sfrmResults"

    ' --- Result entry (one test at a time) ---
    Set frm = StartSingle("SELECT * FROM tblSampleTests ORDER BY DueDate, TestID", "Result Entry", _
        Array("SampleID|Sample|CL|" & SQL_SAMPLES, "MethodID|Method|CL|" & SQL_METHODS, "Status|Status|V|" & VL_TEST_STATUS, _
              "AnalystID|Analyst|C|" & SQL_ANALYSTS, "BatchID|QC Batch|C|" & SQL_BATCHES, _
              "calcWO|Work Order|R|DLookup(""WorkOrderNumber"",""qryHoldTimeStatus"",""TestID="" & Nz([TestID],0))", _
              "PrepDate|Prep Date|D", "AnalysisDate|Analysis Date|D", "DueDate|Due Date|DL", _
              "calcHoldExpires|Hold Expires|R|DLookup(""HoldExpires"",""qryHoldTimeStatus"",""TestID="" & Nz([TestID],0))", _
              "calcHoldStatus|Hold Status|R|DLookup(""HoldStatus"",""qryHoldTimeStatus"",""TestID="" & Nz([TestID],0))"), y)
    frm.AllowAdditions = False
    RedWhen frm.Controls("calcHoldStatus"), "EXPIRED"
    RedWhen frm.Controls("calcHoldStatus"), "OUT OF HOLD"
    y = AddButtonRow(frm, Array( _
        "Mark Reviewed|=ReviewTest([TestID])", _
        "Open Sample|=OpenFormWhere(""frmSamples"",""SampleID"",[SampleID])", _
        "Print COA|=OpenReportWhere(""rptCertificateOfAnalysis"",""WorkOrderID"",DLookup(""WorkOrderID"",""tblSamples"",""SampleID="" & [SampleID]))"), 200, y)
    AddSubform frm, "sfrmResults", "TestID", "TestID", _
        "Results (qualifiers U/J are set automatically from the MDL and RL)", 200, y, 10400, 3600
    FinishObject acForm, frm.Name, "frmResultEntry"

    ' --- QC results subform ---
    Set frm = StartDatasheet("SELECT * FROM tblQCResults ORDER BY QCType, QCResultID", "QC Results", _
        Array("QCType|QC Type|V|" & VL_QC_TYPE, "AnalyteID|Analyte|C|" & SQL_ANALYTES, "ParentSampleID|Parent Sample|C|" & SQL_SAMPLES, _
              "SpikeAmount|Spike|T", "BackgroundValue|Background|T", "MeasuredValue|Measured|T", "DuplicateValue|Duplicate|T", _
              "Units|Units|T", "LowerLimit|Lower Limit|T", "UpperLimit|Upper Limit|T", "RPDLimit|RPD Limit|T", _
              "calcRec|%Rec|R|CalcRecovery([MeasuredValue],[SpikeAmount],[BackgroundValue])", _
              "calcRPD|RPD|R|CalcRPD([MeasuredValue],[DuplicateValue])", _
              "calcQC|QC Status|R|EvalQC([QCType],[MeasuredValue],[SpikeAmount],[BackgroundValue],[DuplicateValue],[LowerLimit],[UpperLimit],[RPDLimit])", _
              "Comments|Comments|T"))
    RedWhen frm.Controls("calcQC"), "FAIL"
    FinishObject acForm, frm.Name, "sfrmQCResults"

    ' --- QC batches ---
    Set frm = StartSingle("SELECT * FROM tblQCBatches ORDER BY BatchDate DESC", "QC Batches", _
        Array("BatchNumber|Batch #|T||""B"" & Format(Now(),""yymmdd-hhnn"")", "MethodID|Method|C|" & SQL_METHODS, _
              "InstrumentID|Instrument|C|" & SQL_INSTRUMENTS, "AnalystID|Analyst|C|" & SQL_ANALYSTS, _
              "BatchDate|Batch Date|D", "Status|Status|V|" & VL_BATCH_STATUS, "Notes|Notes|M"), y)
    AddLabel frm, "Add QC rows for every method analyte:", 200, y, 3400, ROWH
    Set c = CreateControl(frm.Name, acComboBox, acDetail, "", "", 3650, y, 2200, ROWH)
    c.Name = "cboQCType"
    c.RowSourceType = "Value List"
    c.RowSource = VL_QC_TYPE
    c.LimitToList = True
    AddButton frm, "Add QC Rows", "=AddQCRows([BatchID],[cboQCType])", 6000, y - 50, 1800
    AddButton frm, "QC Report", "=OpenReportWhere(""rptQCSummary"",""BatchID"",[BatchID])", 7950, y - 50, 1800
    AddSubform frm, "sfrmQCResults", "BatchID", "BatchID", _
        "QC results (%Rec = (Measured - Background) / Spike x 100;  RPD uses Measured vs Duplicate)", 200, y + 500, 10400, 3600
    FinishObject acForm, frm.Name, "frmQCBatches"
End Sub

' =============================================================================
'  Main menu / dashboard
' =============================================================================

Private Sub AddMenuColumn(ByVal frm As Form, ByVal heading As String, items As Variant, ByVal x As Long, ByVal y As Long)
    Dim i As Long, p() As String
    AddLabel frm, heading, x, y, 3000, 360, 12, True
    y = y + 450
    For i = LBound(items) To UBound(items)
        p = Split(CStr(items(i)), "|")
        AddButton frm, p(0), p(1), x, y, 3000, 420
        y = y + 500
    Next i
End Sub

Private Sub AddKpi(ByVal frm As Form, ByVal cap As String, ByVal expr As String, ByVal queryName As String, ByVal x As Long, ByVal y As Long)
    Dim c As Control
    Set c = CreateControl(frm.Name, acTextBox, acDetail, "", "", x, y, 2000, 620)
    c.ControlSource = "=" & expr
    c.FontSize = 20
    c.FontBold = True
    c.TextAlign = 2
    c.Locked = True
    c.TabStop = False
    c.OnDblClick = "=OpenQueryByName(""" & queryName & """)"
    c.ControlTipText = "Double-click to open the list"
    With AddLabel(frm, cap, x, y + 660, 2000, 300, 9, False)
        .TextAlign = 2
    End With
End Sub

Private Sub BuildMainMenu()
    Dim frm As Form, c As Control, y As Long
    Set frm = NewForm("", "EnviroLIMS", 0)
    frm.RecordSelectors = False
    frm.NavigationButtons = False
    frm.ScrollBars = 0
    frm.Width = 13400
    frm.OnActivate = "=RefreshDashboard()"

    AddLabel frm, "EnviroLIMS", 300, 150, 5000, 600, 24, True
    Set c = CreateControl(frm.Name, acTextBox, acDetail, "", "", 300, 780, 9000, 340)
    c.ControlSource = "=LimsSetting(""LabName"") & ""  -  Environmental Laboratory Information Management System"""
    c.FontSize = 11
    c.Locked = True
    c.TabStop = False
    c.BorderStyle = 0

    y = 1350
    AddKpi frm, "Open tests", "DCount(""*"",""qryOpenTests"")", "qryOpenTests", 300, y
    AddKpi frm, "Overdue tests", "DCount(""*"",""qryOverdueTests"")", "qryOverdueTests", 2450, y
    AddKpi frm, "Hold time at risk", "DCount(""*"",""qryOpenTests"",""HoldStatus In ('EXPIRED','DUE <24H')"")", "qryOpenTests", 4600, y
    AddKpi frm, "Limit exceedances", "DCount(""*"",""qryExceedances"")", "qryExceedances", 6750, y
    AddKpi frm, "QC failures", "DCount(""*"",""qryQCFailures"")", "qryQCFailures", 8900, y
    AddKpi frm, "Calibrations due", "DCount(""*"",""qryInstrumentCalDue"")", "qryInstrumentCalDue", 11050, y

    y = 2900
    AddMenuColumn frm, "Sample Receiving", Array( _
        "Work Orders|=OpenFormByName(""frmWorkOrders"")", _
        "Samples|=OpenFormByName(""frmSamples"")", _
        "Chain of Custody Log|=OpenFormByName(""frmCustodyLog"")", _
        "Work Order Progress|=OpenQueryByName(""qryWorkOrderProgress"")"), 300, y

    AddMenuColumn frm, "Laboratory", Array( _
        "Result Entry (open tests)|=OpenFormFiltered(""frmResultEntry"",""Nz(Status,'') Not In ('Reviewed','Cancelled')"")", _
        "QC Batches|=OpenFormByName(""frmQCBatches"")", _
        "Open Test Backlog|=OpenQueryByName(""qryOpenTests"")", _
        "Overdue Tests|=OpenQueryByName(""qryOverdueTests"")", _
        "Unreviewed Results|=OpenQueryByName(""qryUnreviewedResults"")", _
        "QC Failures|=OpenQueryByName(""qryQCFailures"")", _
        "Calibrations Due|=OpenQueryByName(""qryInstrumentCalDue"")"), 3600, y

    AddMenuColumn frm, "Reports", Array( _
        "Certificate of Analysis...|=PrintCOAPrompt()", _
        "Exceedance Report|=OpenReportByName(""rptExceedances"")", _
        "Holding Time / Backlog|=OpenReportByName(""rptHoldTimes"")", _
        "QC Summary|=OpenReportByName(""rptQCSummary"")", _
        "Chain of Custody|=OpenReportByName(""rptChainOfCustody"")", _
        "Work Order Status|=OpenReportByName(""rptWorkOrderStatus"")", _
        "Location Trends (data)|=OpenQueryByName(""qryLocationTrends"")"), 6900, y

    AddMenuColumn frm, "Setup", Array( _
        "Clients|=OpenFormByName(""frmClients"")", _
        "Projects && Locations|=OpenFormByName(""frmProjects"")", _
        "Methods && Limits|=OpenFormByName(""frmMethods"")", _
        "Analytes|=OpenFormByName(""frmAnalytes"")", _
        "Matrices|=OpenFormByName(""frmMatrices"")", _
        "Regulatory Limits|=OpenFormByName(""frmRegulatoryLimits"")", _
        "Analysts|=OpenFormByName(""frmAnalysts"")", _
        "Instruments|=OpenFormByName(""frmInstruments"")", _
        "Settings|=OpenFormByName(""frmSettings"")", _
        "Audit Log|=OpenFormByName(""frmAuditLog"")"), 10200, y

    AddButton frm, "Refresh Dashboard", "=RefreshDashboard()", 300, 8700, 3000, 420
    frm.Section(acDetail).Height = 9400
    FinishObject acForm, frm.Name, "frmMainMenu"
End Sub

' =============================================================================
'  Reports
'  Column spec: "Caption|FieldOrExpression|WidthInches|Align(L/C/R)|Format|RedValue"
'  Group spec : Array(GroupField, HeaderLine1Expr, HeaderLine2Expr, ...)
' =============================================================================

Private Function AddRptText(ByVal rptName As String, ByVal sec As Integer, ByVal src As String, ByVal x As Long, ByVal y As Long, _
                            ByVal w As Long, ByVal h As Long, ByVal fontSize As Integer, ByVal bold As Boolean, ByVal align As Integer) As Control
    Dim c As Control
    If Left$(src, 1) = "=" Then
        Set c = CreateReportControl(rptName, acTextBox, sec, "", "", x, y, w, h)
        c.ControlSource = src
    Else
        Set c = CreateReportControl(rptName, acTextBox, sec, "", src, x, y, w, h)
    End If
    c.FontSize = fontSize
    c.FontBold = bold
    c.TextAlign = align
    c.BorderStyle = 0
    Set AddRptText = c
End Function

Private Function AddRptLabel(ByVal rptName As String, ByVal sec As Integer, ByVal text As String, ByVal x As Long, ByVal y As Long, _
                             ByVal w As Long, ByVal h As Long, ByVal fontSize As Integer, ByVal bold As Boolean, ByVal align As Integer) As Control
    Dim c As Control
    Set c = CreateReportControl(rptName, acLabel, sec, "", "", x, y, w, h)
    c.Caption = text
    c.FontSize = fontSize
    c.FontBold = bold
    c.TextAlign = align
    Set AddRptLabel = c
End Function

Private Function SpecPart(p() As String, ByVal idx As Long, ByVal defaultValue As String) As String
    If UBound(p) >= idx Then SpecPart = p(idx) Else SpecPart = defaultValue
End Function

Private Function AlignCode(ByVal a As String) As Integer
    Select Case UCase$(a)
        Case "R": AlignCode = 3
        Case "C": AlignCode = 2
        Case Else: AlignCode = 1
    End Select
End Function

Private Sub AddColumnHeaders(ByVal rptName As String, ByVal sec As Integer, cols As Variant, ByVal y As Long)
    Dim i As Long, p() As String, x As Long, w As Long
    x = 0
    For i = LBound(cols) To UBound(cols)
        p = Split(CStr(cols(i)), "|")
        w = CLng(Val(p(2)) * 1440)
        AddRptLabel rptName, sec, p(0), x, y, w - 40, 260, 8, True, AlignCode(SpecPart(p, 3, "L"))
        x = x + w
    Next i
    CreateReportControl rptName, acLine, sec, "", "", 0, y + 290, RPT_W, 0
End Sub

Private Sub EnsurePageSections(ByVal rpt As Report)
    Dim h As Long
    On Error Resume Next
    h = rpt.Section(acPageHeader).Height
    If Err.Number <> 0 Then
        Err.Clear
        DoCmd.RunCommand acCmdPageHdrFtr
    End If
End Sub

Private Sub BuildReport(ByVal rptName As String, ByVal src As String, ByVal title As String, groups As Variant, sorts As Variant, _
                        cols As Variant, Optional ByVal footerNote As String = "", Optional ByVal newPagePerGroup As Boolean = False)
    Dim rpt As Report, tmp As String, i As Long, j As Long, g As Variant, sec As Integer, y As Long, x As Long, w As Long
    Dim nGroups As Long, p() As String, c As Control

    Set rpt = CreateReport()
    tmp = rpt.Name
    rpt.RecordSource = src
    rpt.Caption = title
    rpt.Width = RPT_W
    EnsurePageSections rpt

    ' Page header
    AddRptLabel tmp, acPageHeader, title, 0, 0, 6000, 450, 16, True, 1
    AddRptText tmp, acPageHeader, "=LimsSetting(""LabName"")", 5000, 0, RPT_W - 5000, 300, 11, True, 3
    AddRptText tmp, acPageHeader, "=LimsSetting(""LabAddress"") & ""   "" & LimsSetting(""LabPhone"")", 3000, 300, RPT_W - 3000, 260, 8, False, 3
    y = 650
    If IsArray(groups) Then nGroups = UBound(groups) - LBound(groups) + 1
    If nGroups = 0 Then
        AddColumnHeaders tmp, acPageHeader, cols, y
        y = y + 340
    End If
    rpt.Section(acPageHeader).Height = y + 60

    ' Group headers (the innermost one carries the column headings)
    For i = 0 To nGroups - 1
        g = groups(LBound(groups) + i)
        CreateGroupLevel tmp, CStr(g(0)), True, False
        sec = 5 + i * 2
        y = 80
        For j = 1 To UBound(g)
            AddRptText tmp, sec, CStr(g(j)), 0, y, RPT_W, 290, IIf(j = 1 And i = 0, 11, 9), (j = 1), 1
            y = y + 300
        Next j
        If i = nGroups - 1 Then
            y = y + 80
            AddColumnHeaders tmp, sec, cols, y
            y = y + 340
        End If
        rpt.Section(sec).Height = y + 40
        rpt.GroupLevel(i).KeepTogether = 2
    Next i
    If newPagePerGroup And nGroups > 0 Then rpt.Section(5).ForceNewPage = 1

    ' Sorting within groups
    If IsArray(sorts) Then
        For i = LBound(sorts) To UBound(sorts)
            CreateGroupLevel tmp, CStr(sorts(i)), False, False
        Next i
    End If

    ' Detail
    x = 0
    For i = LBound(cols) To UBound(cols)
        p = Split(CStr(cols(i)), "|")
        w = CLng(Val(p(2)) * 1440)
        Set c = AddRptText(tmp, acDetail, p(1), x, 20, w - 40, 260, 8, False, AlignCode(SpecPart(p, 3, "L")))
        c.CanGrow = True
        If UBound(p) >= 4 Then
            If Len(p(4)) > 0 Then c.Format = p(4)
        End If
        If UBound(p) >= 5 Then
            If Len(p(5)) > 0 Then RedWhen c, p(5)
        End If
        x = x + w
    Next i
    rpt.Section(acDetail).Height = 300

    ' Page footer
    CreateReportControl tmp, acLine, acPageFooter, "", "", 0, 40, RPT_W, 0
    AddRptText tmp, acPageFooter, "=""Printed "" & Format(Now(),""mm/dd/yyyy hh:nn"")", 0, 80, 3000, 260, 8, False, 1
    AddRptText tmp, acPageFooter, "=""Page "" & [Page] & "" of "" & [Pages]", RPT_W - 2500, 80, 2500, 260, 8, False, 3
    If Len(footerNote) > 0 Then
        If Left$(footerNote, 1) <> "=" Then footerNote = "=""" & Replace(footerNote, """", """""") & """"
        AddRptText tmp, acPageFooter, footerNote, 0, 360, RPT_W, 480, 7, False, 1
        rpt.Section(acPageFooter).Height = 900
    Else
        rpt.Section(acPageFooter).Height = 400
    End If

    FinishObject acReport, tmp, rptName
End Sub

Private Sub BuildReports()
    ' --- Certificate of Analysis ---
    BuildReport "rptCertificateOfAnalysis", "qryResultsWithLimits", "Certificate of Analysis", _
        Array( _
            Array("WorkOrderNumber", _
                  "=""Work Order: "" & [WorkOrderNumber] & ""      Client: "" & [ClientName]", _
                  "=""Project: "" & Nz([ProjectName],""-"") & ""      Received: "" & Format([ReceivedDate],""mm/dd/yyyy hh:nn"") & ""      Limit set: "" & Nz([LimitSet],""None"")"), _
            Array("LabSampleID", _
                  "=""Lab Sample ID: "" & [LabSampleID] & ""      Client Sample ID: "" & Nz([ClientSampleID],"""") & ""      Location: "" & Nz([LocationCode],""-"")", _
                  "=""Matrix: "" & [MatrixName] & ""      Type: "" & Nz([SampleType],"""") & ""      Collected: "" & Format([CollectedDate],""mm/dd/yyyy hh:nn"")")), _
        Array("MethodCode", "AnalyteName"), _
        Array("Method|MethodCode|0.95|L", "Analyte|AnalyteName|1.5|L", "Result|ReportedResult|0.8|R", "Qual|Qualifier|0.45|C", _
              "Units|Units|0.75|L", "DF|DilutionFactor|0.4|R", "MDL|MDL|0.65|R", "RL|RL|0.65|R", "Limit|LimitValue|0.65|R", _
              "Flag|ExceedFlag|0.7|C||EXCEEDS"), _
        "=""U = not detected above the MDL (reported as ND).  J = estimated value between the MDL and RL.  EXCEEDS = above the limit for the project's limit set.  " & _
        "Results relate only to the samples tested.     Approved by: "" & LimsSetting(""LabDirector"") & ""     "" & LimsSetting(""Accreditation"")", True

    ' --- Exceedances ---
    BuildReport "rptExceedances", "qryExceedances", "Regulatory Limit Exceedances", Empty, Array("CollectedDate", "LabSampleID"), _
        Array("Work Order|WorkOrderNumber|1.0|L", "Lab ID|LabSampleID|1.0|L", "Location|LocationCode|0.8|L", _
              "Collected|CollectedDate|0.9|L|mm/dd/yy hh:nn", "Analyte|AnalyteName|1.3|L", "Result|ReportedResult|0.7|R", _
              "Units|Units|0.5|L", "Limit|LimitValue|0.6|R", "Limit Set|LimitSet|0.7|L")

    ' --- Holding times / backlog ---
    BuildReport "rptHoldTimes", "qryOpenTests", "Open Tests - Holding Time Status", Empty, Array("HoldExpires", "DueDate"), _
        Array("Work Order|WorkOrderNumber|1.0|L", "Lab ID|LabSampleID|1.0|L", "Method|MethodCode|0.9|L", _
              "Collected|CollectedDate|1.0|L|mm/dd/yy hh:nn", "Hold Expires|HoldExpires|1.0|L|mm/dd/yy hh:nn", _
              "Hrs Left|HoursRemaining|0.6|R", "Hold|HoldStatus|0.8|C||EXPIRED", "Due|DueDate|0.6|L|mm/dd/yy", "Status|Status|0.6|L")

    ' --- QC summary ---
    BuildReport "rptQCSummary", "qryQCEvaluation", "Quality Control Summary", _
        Array(Array("BatchNumber", "=""Batch: "" & [BatchNumber] & ""      Method: "" & [MethodCode] & ""      Date: "" & Format([BatchDate],""mm/dd/yyyy"")")), _
        Array("QCType", "AnalyteName"), _
        Array("QC Type|QCType|1.0|L", "Analyte|AnalyteName|1.4|L", "Spike|SpikeAmount|0.65|R", "Bkgd|BackgroundValue|0.65|R", _
              "Result|MeasuredValue|0.65|R", "Dup|DuplicateValue|0.65|R", "%Rec|PctRecovery|0.55|R", "RPD|RPD|0.5|R", _
              "Limits|=Nz([LowerLimit],"""") & "" - "" & Nz([UpperLimit],"""")|0.7|C", "Status|QCStatus|0.75|C||FAIL")

    ' --- Chain of custody ---
    BuildReport "rptChainOfCustody", "qryCustodyReport", "Chain of Custody", _
        Array(Array("WorkOrderNumber", "=""Work Order: "" & [WorkOrderNumber] & ""      Client: "" & [ClientName]")), _
        Array("EventDate"), _
        Array("Date/Time|EventDate|1.2|L|mm/dd/yyyy hh:nn", "Action|CustodyAction|1.5|L", "Relinquished By|RelinquishedBy|1.2|L", _
              "Received By|ReceivedBy|1.2|L", "Location|StorageLocation|1.1|L", "Comments|Comments|1.3|L"), "", True

    ' --- Work order status ---
    BuildReport "rptWorkOrderStatus", "qryWorkOrderProgress", "Work Order Status", Empty, Array("WorkOrderNumber"), _
        Array("Work Order|WorkOrderNumber|1.1|L", "Client|ClientName|2.0|L", "Received|ReceivedDate|0.9|L|mm/dd/yy", _
              "Due|DueDate|0.8|L|mm/dd/yy", "Status|Status|0.9|L", "Tests|TotalTests|0.5|R", "Reviewed|ReviewedTests|0.6|R", _
              "Complete|PctComplete|0.7|R|0%")
End Sub
