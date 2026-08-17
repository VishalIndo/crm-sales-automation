Attribute VB_Name = "modSchema"
Option Explicit

Public Sub SetupCRMApp()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    BuildListsSheet
    BuildSettingsSheet
    BuildLeadsSheet
    BuildActivitiesSheet
    BuildQuotesSheet
    BuildDashboardSheet
    BuildHomeSheet
    InstallHomeButtons
    RefreshDashboard

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "CRM workbook structure is ready.", vbInformation, APP_TITLE
End Sub

Public Sub Auto_Open()
    If WorksheetExists(SHEET_LEADS) Then
        InitializeAppOnOpen
    End If
End Sub

Public Sub InitializeAppOnOpen()
    HighlightFollowUps
    ShowTodayFollowUps False
End Sub

Private Sub BuildListsSheet()
    Dim ws As Worksheet
    Dim i As Long
    Dim values As Variant

    Set ws = EnsureSheet(SHEET_LISTS)
    ws.Cells.Clear

    ws.Range("A1").Value = "Stages"
    values = StageList()
    For i = LBound(values) To UBound(values)
        ws.Cells(i + 2, 1).Value = values(i)
    Next i

    ws.Range("B1").Value = "LeadSources"
    values = SourceList()
    For i = LBound(values) To UBound(values)
        ws.Cells(i + 2, 2).Value = values(i)
    Next i

    ws.Range("C1").Value = "ActivityTypes"
    values = ActivityTypeList()
    For i = LBound(values) To UBound(values)
        ws.Cells(i + 2, 3).Value = values(i)
    Next i

    ws.Range("D1").Value = "QuoteStatuses"
    values = QuoteStatusList()
    For i = LBound(values) To UBound(values)
        ws.Cells(i + 2, 4).Value = values(i)
    Next i

    ws.Range("E1").Value = "Statuses"
    values = StatusList()
    For i = LBound(values) To UBound(values)
        ws.Cells(i + 2, 5).Value = values(i)
    Next i

    ws.Rows(1).Font.Bold = True
    AutoFitUsedColumns ws
End Sub

Private Sub BuildSettingsSheet()
    Dim ws As Worksheet

    Set ws = EnsureSheet(SHEET_SETTINGS)
    ws.Cells.Clear

    ws.Range("A1").Value = "Setting"
    ws.Range("B1").Value = "Value"
    ws.Range("A2").Value = "DefaultOwner"
    ws.Range("B2").Value = Environ$("Username")
    ws.Range("A3").Value = "FollowUpAlertDays"
    ws.Range("B3").Value = 0
    ws.Rows(1).Font.Bold = True
    AutoFitUsedColumns ws
End Sub

Private Sub BuildLeadsSheet()
    Dim ws As Worksheet
    Dim tableObj As ListObject

    Set ws = EnsureSheet(SHEET_LEADS)
    ws.Cells.Clear
    Set tableObj = EnsureTable(ws, TABLE_LEADS, LeadHeaders())

    ApplyValidation tableObj.ListColumns("Stage").DataBodyRange, JoinArray(StageList())
    ApplyValidation tableObj.ListColumns("LeadSource").DataBodyRange, JoinArray(SourceList())
    ApplyValidation tableObj.ListColumns("Status").DataBodyRange, JoinArray(StatusList())
    ApplyDateFormat tableObj.ListColumns("CreatedDate").Range
    ApplyDateFormat tableObj.ListColumns("NextFollowUp").Range
    ApplyDateFormat tableObj.ListColumns("LastContactDate").Range
    ApplyMoneyFormat tableObj.ListColumns("ExpectedValue").DataBodyRange
    AutoFitUsedColumns ws
End Sub

Private Sub BuildActivitiesSheet()
    Dim ws As Worksheet
    Dim tableObj As ListObject

    Set ws = EnsureSheet(SHEET_ACTIVITIES)
    ws.Cells.Clear
    Set tableObj = EnsureTable(ws, TABLE_ACTIVITIES, ActivityHeaders())

    ApplyValidation tableObj.ListColumns("ActivityType").DataBodyRange, JoinArray(ActivityTypeList())
    ApplyDateFormat tableObj.ListColumns("ActivityDate").Range
    ApplyDateFormat tableObj.ListColumns("NextFollowUp").Range
    AutoFitUsedColumns ws
End Sub

Private Sub BuildQuotesSheet()
    Dim ws As Worksheet
    Dim tableObj As ListObject

    Set ws = EnsureSheet(SHEET_QUOTES)
    ws.Cells.Clear
    Set tableObj = EnsureTable(ws, TABLE_QUOTES, QuoteHeaders())

    ApplyValidation tableObj.ListColumns("QuoteStatus").DataBodyRange, JoinArray(QuoteStatusList())
    ApplyDateFormat tableObj.ListColumns("QuoteDate").Range
    ApplyDateFormat tableObj.ListColumns("ValidUntil").Range
    ApplyMoneyFormat tableObj.ListColumns("Amount").DataBodyRange
    AutoFitUsedColumns ws
End Sub

Private Sub BuildDashboardSheet()
    Dim ws As Worksheet

    Set ws = EnsureSheet(SHEET_DASHBOARD)
    ws.Cells.Clear
    DeleteDecorations ws
    ws.Cells.Interior.Color = RGB(245, 247, 250)
    ActiveWindow.DisplayGridlines = False

    With ws.Range("A1:J2")
        .Merge
        .Value = APP_COMPANY & " | Sales Dashboard"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Size = 18
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(104, 118, 143)
    End With

    BuildMetricCard ws, "B4:C4", "B5:C6", "TOTAL LEADS", RGB(56, 89, 138)
    BuildMetricCard ws, "D4:E4", "D5:E6", "OPEN LEADS", RGB(214, 51, 108)
    BuildMetricCard ws, "F4:G4", "F5:G6", "FOLLOW-UPS", RGB(42, 157, 143)
    BuildMetricCard ws, "H4:I4", "H5:I6", "QUOTE VALUE", RGB(244, 162, 97)

    ws.Range("B8").Value = "PIPELINE SUMMARY"
    ws.Range("B8:E8").Merge
    ws.Range("B8:E8").HorizontalAlignment = xlCenter
    ws.Range("B8:E8").Interior.Color = RGB(222, 226, 230)
    ws.Range("B8:E8").Font.Bold = True

    ws.Range("B9").Value = "Stage"
    ws.Range("C9").Value = "Count"
    ws.Range("D9").Value = "Won"
    ws.Range("E9").Value = "Lost"
    ws.Range("B9:E9").Font.Bold = True
    ws.Range("B9:E9").Interior.Color = RGB(235, 238, 242)

    ws.Range("G8").Value = "TEAM NOTE"
    ws.Range("G8:I8").Merge
    ws.Range("G8:I8").HorizontalAlignment = xlCenter
    ws.Range("G8:I8").Interior.Color = RGB(222, 226, 230)
    ws.Range("G8:I8").Font.Bold = True

    ws.Range("G9:I14").Merge
    ws.Range("G9:I14").Value = "Use this dashboard as the management view." & vbCrLf & vbCrLf & _
        "1. Review leads due today." & vbCrLf & _
        "2. Track open pipeline count." & vbCrLf & _
        "3. Check quotation value." & vbCrLf & _
        "4. Refresh after updates."
    ws.Range("G9:I14").WrapText = True
    ws.Range("G9:I14").VerticalAlignment = xlTop
    ws.Range("G9:I14").Interior.Color = RGB(250, 250, 250)
    ws.Range("G9:I14").Borders.LineStyle = xlContinuous

    ws.Range("B18").Value = "MONTHLY SUMMARY"
    ws.Range("B18:E18").Merge
    ws.Range("B18:E18").HorizontalAlignment = xlCenter
    ws.Range("B18:E18").Interior.Color = RGB(222, 226, 230)
    ws.Range("B18:E18").Font.Bold = True

    ws.Range("B19").Value = "Won Leads"
    ws.Range("B20").Value = "Lost Leads"
    ws.Range("B21").Value = "Quotes This Month"
    ws.Range("B22").Value = "Open Quote Value"
    ws.Range("C19:C22").Interior.Color = RGB(250, 250, 250)
    ws.Range("B19:C22").Borders.LineStyle = xlContinuous

    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B:I").ColumnWidth = 14
    AutoFitUsedColumns ws
End Sub

Private Sub BuildHomeSheet()
    Dim ws As Worksheet

    Set ws = EnsureSheet(SHEET_HOME)
    ws.Cells.Clear
    DeleteDecorations ws
    ws.Cells.Interior.Color = RGB(245, 247, 250)
    ActiveWindow.DisplayGridlines = False

    ws.Cells.Font.Name = "Calibri"
    ws.Rows("1:30").RowHeight = 24
    ws.Columns("A").ColumnWidth = 4
    ws.Columns("B").ColumnWidth = 16
    ws.Columns("C").ColumnWidth = 14
    ws.Columns("D").ColumnWidth = 14
    ws.Columns("E").ColumnWidth = 14
    ws.Columns("F").ColumnWidth = 14
    ws.Columns("G").ColumnWidth = 14
    ws.Columns("H").ColumnWidth = 14
    ws.Columns("I").ColumnWidth = 14
    ws.Columns("J").ColumnWidth = 14
    ws.Columns("K").ColumnWidth = 14

    With ws.Range("A1:K2")
        .Merge
        .Value = APP_COMPANY & " CRM Workspace"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Size = 20
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(104, 118, 143)
    End With

    CreatePanel ws, "B4:C26", RGB(57, 72, 103), ""
    CreatePanel ws, "D4:K8", RGB(255, 255, 255), ""
    CreatePanel ws, "D10:F17", RGB(255, 255, 255), ""
    CreatePanel ws, "G10:I17", RGB(255, 255, 255), ""
    CreatePanel ws, "J10:K17", RGB(255, 255, 255), ""
    CreatePanel ws, "D19:F26", RGB(255, 255, 255), ""
    CreatePanel ws, "G19:I26", RGB(255, 255, 255), ""
    CreatePanel ws, "J19:K26", RGB(255, 255, 255), ""

    With ws.Range("B5:C6")
        .Merge
        .Value = APP_COMPANY
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
    End With

    With ws.Range("B8:C9")
        .Merge
        .Value = "NAVIGATION"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
    End With

    With ws.Range("D5:H5")
        .Merge
        .Value = "Sales Control Center"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 16
    End With

    With ws.Range("D6:K7")
        .Merge
        .Value = "Manage leads, activities, quotations, follow-ups, and reporting from one workbook."
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Font.Size = 11
        .Font.Color = RGB(95, 99, 104)
    End With

    WritePanelHeader ws, "D10:F10", "Lead Management", RGB(56, 89, 138)
    WritePanelText ws, "D11:F12", "Create new leads, search records, and keep the pipeline current."

    WritePanelHeader ws, "G10:I10", "Activity Tracking", RGB(42, 157, 143)
    WritePanelText ws, "G11:I12", "Record calls, meetings, and follow-up commitments for each prospect."

    WritePanelHeader ws, "J10:K10", "Quotations", RGB(244, 162, 97)
    WritePanelText ws, "J11:K12", "Create and track commercial offers linked to active leads."

    WritePanelHeader ws, "D19:F19", "Pipeline Actions", RGB(214, 51, 108)
    WritePanelText ws, "D20:F21", "Advance stage, review due work, and keep momentum on open deals."

    WritePanelHeader ws, "G19:I19", "Reporting", RGB(98, 111, 164)
    WritePanelText ws, "G20:I21", "Refresh the dashboard after updates and review the management summary."

    WritePanelHeader ws, "J19:K19", "Pages", RGB(120, 128, 145)
    WritePanelText ws, "J20:K24", "Home" & vbCrLf & "Leads" & vbCrLf & "Activities" & vbCrLf & "Quotations" & vbCrLf & "Dashboard"
End Sub

Private Sub InstallHomeButtons()
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets(SHEET_HOME)
    DeleteExistingButtons ws

    CreateNavButton ws, "Home", "", ws.Range("B11:C12"), RGB(126, 138, 171), False
    CreateNavButton ws, "Leads", "GoToLeads", ws.Range("B13:C14"), RGB(56, 89, 138), True
    CreateNavButton ws, "Activities", "", ws.Range("B15:C16"), RGB(42, 157, 143), False
    CreateNavButton ws, "Quotes", "", ws.Range("B17:C18"), RGB(244, 162, 97), False
    CreateNavButton ws, "Dashboard", "RefreshDashboard", ws.Range("B19:C20"), RGB(98, 111, 164), True

    CreateActionButton ws, "Add Lead", "AddLead", ws.Range("D13:F15"), RGB(56, 89, 138)
    CreateActionButton ws, "Search Lead", "SearchLead", ws.Range("D15:F17"), RGB(87, 116, 164)

    CreateActionButton ws, "Add Activity", "AddActivityForSelectedLead", ws.Range("G13:I15"), RGB(42, 157, 143)
    CreateActionButton ws, "Show Follow-Ups", "ShowTodayFollowUpsManual", ws.Range("G15:I17"), RGB(66, 181, 167)

    CreateActionButton ws, "Add Quote", "AddQuotationForSelectedLead", ws.Range("J13:K15"), RGB(244, 162, 97)
    CreateActionButton ws, "Advance Stage", "AdvanceSelectedLeadStage", ws.Range("D22:F24"), RGB(214, 51, 108)
    CreateActionButton ws, "Refresh Dashboard", "RefreshDashboard", ws.Range("G22:I24"), RGB(98, 111, 164)
    CreateActionButton ws, "Go To Leads", "GoToLeads", ws.Range("J22:K24"), RGB(120, 128, 145)
End Sub

Private Sub CreateActionButton(ByVal ws As Worksheet, ByVal caption As String, ByVal macroName As String, ByVal targetRange As Range, ByVal fillColor As Long)
    Dim shp As Shape

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, targetRange.Left + 8, targetRange.Top + 8, targetRange.Width - 16, targetRange.Height - 16)
    shp.Name = BUTTON_PREFIX & Replace(caption, " ", "")
    shp.TextFrame.Characters.Text = caption
    shp.Fill.ForeColor.RGB = fillColor
    shp.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    shp.TextFrame.Characters.Font.Size = 13
    shp.TextFrame.Characters.Font.Bold = True
    shp.Line.ForeColor.RGB = RGB(18, 43, 69)
    shp.Shadow.Visible = msoTrue
    shp.OnAction = macroName
End Sub

Private Sub CreateNavButton(ByVal ws As Worksheet, ByVal caption As String, ByVal macroName As String, ByVal targetRange As Range, ByVal fillColor As Long, ByVal isClickable As Boolean)
    Dim shp As Shape

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, targetRange.Left + 6, targetRange.Top + 4, targetRange.Width - 12, targetRange.Height - 8)
    shp.Name = BUTTON_PREFIX & "Nav" & Replace(caption, " ", "")
    shp.TextFrame.Characters.Text = caption
    shp.Fill.ForeColor.RGB = fillColor
    shp.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    shp.TextFrame.Characters.Font.Size = 11
    shp.TextFrame.Characters.Font.Bold = True
    shp.Line.Visible = msoFalse
    If isClickable Then
        shp.OnAction = macroName
    End If
End Sub

Private Sub DeleteExistingButtons(ByVal ws As Worksheet)
    Dim i As Long

    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, Len(BUTTON_PREFIX)) = BUTTON_PREFIX _
            Or Left$(ws.Shapes(i).Name, Len(DECOR_PREFIX)) = DECOR_PREFIX Then
            ws.Shapes(i).Delete
        End If
    Next i
End Sub

Private Sub DeleteDecorations(ByVal ws As Worksheet)
    DeleteExistingButtons ws
End Sub

Private Sub CreatePanel(ByVal ws As Worksheet, ByVal addressText As String, ByVal fillColor As Long, ByVal labelText As String)
    Dim rng As Range
    Dim shp As Shape

    Set rng = ws.Range(addressText)
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, rng.Left, rng.Top, rng.Width, rng.Height)
    shp.Name = DECOR_PREFIX & Replace(addressText, ":", "_")
    shp.Fill.ForeColor.RGB = fillColor
    shp.Line.ForeColor.RGB = RGB(218, 223, 230)
    shp.Shadow.Visible = msoTrue
    If Len(labelText) > 0 Then
        shp.TextFrame.Characters.Text = labelText
    End If
    shp.ZOrder msoSendToBack
End Sub

Private Sub BuildMetricCard(ByVal ws As Worksheet, ByVal titleAddress As String, ByVal valueAddress As String, ByVal titleText As String, ByVal fillColor As Long)
    With ws.Range(titleAddress)
        .Merge
        .Value = titleText
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = fillColor
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
    End With

    With ws.Range(valueAddress)
        .Merge
        .Value = 0
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = fillColor
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 16
        .Borders.LineStyle = xlContinuous
    End With
End Sub

Private Sub WritePanelHeader(ByVal ws As Worksheet, ByVal addressText As String, ByVal titleText As String, ByVal fillColor As Long)
    With ws.Range(addressText)
        .Merge
        .Value = titleText
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 12
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = fillColor
    End With
End Sub

Private Sub WritePanelText(ByVal ws As Worksheet, ByVal addressText As String, ByVal textValue As String)
    With ws.Range(addressText)
        .Merge
        .Value = textValue
        .WrapText = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .Font.Size = 10
        .Font.Color = RGB(70, 74, 80)
    End With
End Sub

Private Sub ApplyValidation(ByVal rng As Range, ByVal csvValues As String)
    On Error Resume Next
    rng.Validation.Delete
    On Error GoTo 0

    rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=csvValues
    rng.Validation.InCellDropdown = True
End Sub
