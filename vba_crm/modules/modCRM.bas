Attribute VB_Name = "modCRM"
Option Explicit

Public Sub GoToLeads()
    ThisWorkbook.Worksheets(SHEET_LEADS).Activate
    ThisWorkbook.Worksheets(SHEET_LEADS).Range("A2").Select
End Sub

Public Sub AddLead()
    Dim ws As Worksheet
    Dim nextRow As Long
    Dim companyName As String
    Dim contactName As String
    Dim phoneValue As String
    Dim emailValue As String
    Dim sourceValue As String
    Dim industryValue As String
    Dim ownerValue As String
    Dim stageValue As String
    Dim expectedValue As Variant
    Dim probabilityValue As Variant
    Dim followUpValue As Variant
    Dim notesValue As String

    companyName = PromptText(APP_TITLE, "Company name:")
    If Len(companyName) = 0 Then Exit Sub

    contactName = PromptText(APP_TITLE, "Contact name:")
    phoneValue = PromptText(APP_TITLE, "Phone:")
    emailValue = PromptText(APP_TITLE, "Email:")
    sourceValue = PromptText(APP_TITLE, "Lead source:", "Referral")
    industryValue = PromptText(APP_TITLE, "Industry:")
    ownerValue = PromptText(APP_TITLE, "Assigned to:", DefaultOwner())
    stageValue = PromptText(APP_TITLE, "Stage:", "New")
    expectedValue = PromptNumberValue(APP_TITLE, "Expected value:", "0")
    If IsEmpty(expectedValue) Then expectedValue = 0
    probabilityValue = PromptNumberValue(APP_TITLE, "Probability percent:", "10")
    If IsEmpty(probabilityValue) Then probabilityValue = 10
    followUpValue = PromptDateValue(APP_TITLE, "Next follow-up date:", Format(Date + 3, "yyyy-mm-dd"))
    notesValue = PromptText(APP_TITLE, "Notes:")

    Set ws = ThisWorkbook.Worksheets(SHEET_LEADS)
    nextRow = LastUsedRow(ws, 1) + 1

    ws.Cells(nextRow, 1).Value = NextId("LD", ws, 1)
    ws.Cells(nextRow, 2).Value = Date
    ws.Cells(nextRow, 3).Value = companyName
    ws.Cells(nextRow, 4).Value = contactName
    ws.Cells(nextRow, 5).Value = phoneValue
    ws.Cells(nextRow, 6).Value = emailValue
    ws.Cells(nextRow, 7).Value = sourceValue
    ws.Cells(nextRow, 8).Value = industryValue
    ws.Cells(nextRow, 9).Value = ownerValue
    ws.Cells(nextRow, 10).Value = stageValue
    ws.Cells(nextRow, 11).Value = expectedValue
    ws.Cells(nextRow, 12).Value = probabilityValue
    ws.Cells(nextRow, 13).Value = followUpValue
    ws.Cells(nextRow, 14).Value = "Open"
    ws.Cells(nextRow, 15).Value = Date
    ws.Cells(nextRow, 16).Value = notesValue

    HighlightFollowUps
    RefreshDashboard
    ws.Activate
    ws.Cells(nextRow, 1).Select

    MsgBox "Lead added: " & ws.Cells(nextRow, 1).Value, vbInformation, APP_TITLE
End Sub

Public Sub SearchLead()
    Dim ws As Worksheet
    Dim term As String
    Dim lastRow As Long
    Dim rowIndex As Long

    term = LCase$(PromptText(APP_TITLE, "Search by company, contact, email, or phone:"))
    If Len(term) = 0 Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(SHEET_LEADS)
    lastRow = LastUsedRow(ws, 1)

    For rowIndex = 2 To lastRow
        If InStr(1, LCase$(SafeText(ws.Cells(rowIndex, 3).Value)), term, vbTextCompare) > 0 _
            Or InStr(1, LCase$(SafeText(ws.Cells(rowIndex, 4).Value)), term, vbTextCompare) > 0 _
            Or InStr(1, LCase$(SafeText(ws.Cells(rowIndex, 6).Value)), term, vbTextCompare) > 0 _
            Or InStr(1, LCase$(SafeText(ws.Cells(rowIndex, 5).Value)), term, vbTextCompare) > 0 Then
            ws.Activate
            ws.Cells(rowIndex, 1).Select
            MsgBox "Lead found: " & ws.Cells(rowIndex, 1).Value, vbInformation, APP_TITLE
            Exit Sub
        End If
    Next rowIndex

    MsgBox "No lead matched your search.", vbExclamation, APP_TITLE
End Sub

Public Sub AddActivityForSelectedLead()
    Dim leadRow As Long
    Dim leadsWs As Worksheet
    Dim activityWs As Worksheet
    Dim nextRow As Long
    Dim leadId As String
    Dim activityType As String
    Dim summaryText As String
    Dim nextAction As String
    Dim nextFollowUp As Variant
    Dim ownerValue As String
    Dim outcomeValue As String

    leadRow = SelectedLeadDataRow()
    If leadRow = 0 Then Exit Sub

    Set leadsWs = ThisWorkbook.Worksheets(SHEET_LEADS)
    Set activityWs = ThisWorkbook.Worksheets(SHEET_ACTIVITIES)
    leadId = GetLeadIdFromRow(leadRow)

    activityType = PromptText(APP_TITLE, "Activity type:", "Call")
    If Len(activityType) = 0 Then Exit Sub

    summaryText = PromptText(APP_TITLE, "Activity summary:")
    If Len(summaryText) = 0 Then Exit Sub

    nextAction = PromptText(APP_TITLE, "Next action:")
    nextFollowUp = PromptDateValue(APP_TITLE, "Next follow-up date:", Format(Date + 3, "yyyy-mm-dd"))
    ownerValue = PromptText(APP_TITLE, "Owner:", SafeText(leadsWs.Cells(leadRow, 9).Value))
    outcomeValue = PromptText(APP_TITLE, "Outcome:", "Pending")

    nextRow = LastUsedRow(activityWs, 1) + 1
    activityWs.Cells(nextRow, 1).Value = NextId("AC", activityWs, 1)
    activityWs.Cells(nextRow, 2).Value = leadId
    activityWs.Cells(nextRow, 3).Value = Date
    activityWs.Cells(nextRow, 4).Value = activityType
    activityWs.Cells(nextRow, 5).Value = summaryText
    activityWs.Cells(nextRow, 6).Value = nextAction
    activityWs.Cells(nextRow, 7).Value = nextFollowUp
    activityWs.Cells(nextRow, 8).Value = ownerValue
    activityWs.Cells(nextRow, 9).Value = outcomeValue

    leadsWs.Cells(leadRow, 13).Value = nextFollowUp
    leadsWs.Cells(leadRow, 15).Value = Date

    HighlightFollowUps
    RefreshDashboard
    MsgBox "Activity saved for " & leadId, vbInformation, APP_TITLE
End Sub

Public Sub AddQuotationForSelectedLead()
    Dim leadRow As Long
    Dim quotesWs As Worksheet
    Dim nextRow As Long
    Dim leadId As String
    Dim amountValue As Variant
    Dim quoteStatus As String
    Dim validUntil As Variant
    Dim notesValue As String

    leadRow = SelectedLeadDataRow()
    If leadRow = 0 Then Exit Sub

    leadId = GetLeadIdFromRow(leadRow)
    amountValue = PromptNumberValue(APP_TITLE, "Quotation amount:", "0")
    If IsEmpty(amountValue) Then Exit Sub

    quoteStatus = PromptText(APP_TITLE, "Quote status:", "Sent")
    If Len(quoteStatus) = 0 Then Exit Sub

    validUntil = PromptDateValue(APP_TITLE, "Valid until:", Format(Date + 14, "yyyy-mm-dd"))
    notesValue = PromptText(APP_TITLE, "Quote notes:")

    Set quotesWs = ThisWorkbook.Worksheets(SHEET_QUOTES)
    nextRow = LastUsedRow(quotesWs, 1) + 1

    quotesWs.Cells(nextRow, 1).Value = NextId("QT", quotesWs, 1)
    quotesWs.Cells(nextRow, 2).Value = leadId
    quotesWs.Cells(nextRow, 3).Value = Date
    quotesWs.Cells(nextRow, 4).Value = amountValue
    quotesWs.Cells(nextRow, 5).Value = quoteStatus
    quotesWs.Cells(nextRow, 6).Value = validUntil
    quotesWs.Cells(nextRow, 7).Value = notesValue

    ThisWorkbook.Worksheets(SHEET_LEADS).Cells(leadRow, 10).Value = "Quoted"

    RefreshDashboard
    MsgBox "Quotation saved for " & leadId, vbInformation, APP_TITLE
End Sub

Public Sub AdvanceSelectedLeadStage()
    Dim leadRow As Long
    Dim ws As Worksheet
    Dim currentStage As String
    Dim nextStage As String

    leadRow = SelectedLeadDataRow()
    If leadRow = 0 Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(SHEET_LEADS)
    currentStage = SafeText(ws.Cells(leadRow, 10).Value)
    nextStage = NextStageValue(currentStage)

    If Len(nextStage) = 0 Then
        MsgBox "No next stage found for " & currentStage, vbInformation, APP_TITLE
        Exit Sub
    End If

    ws.Cells(leadRow, 10).Value = nextStage
    If nextStage = "Won" Or nextStage = "Lost" Then
        ws.Cells(leadRow, 14).Value = "Closed"
    End If

    RefreshDashboard
    MsgBox "Lead moved to stage: " & nextStage, vbInformation, APP_TITLE
End Sub

Private Function NextStageValue(ByVal currentStage As String) As String
    Dim stages As Variant
    Dim i As Long

    stages = StageList()
    For i = LBound(stages) To UBound(stages) - 1
        If StrComp(CStr(stages(i)), currentStage, vbTextCompare) = 0 Then
            NextStageValue = CStr(stages(i + 1))
            Exit Function
        End If
    Next i

    If Len(currentStage) = 0 Then
        NextStageValue = CStr(stages(LBound(stages)))
    End If
End Function

Public Sub HighlightFollowUps()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim followUpDate As Variant
    Dim statusValue As String

    Set ws = ThisWorkbook.Worksheets(SHEET_LEADS)
    lastRow = LastUsedRow(ws, 1)

    For rowIndex = 2 To lastRow
        ws.Rows(rowIndex).Interior.Pattern = xlNone
        followUpDate = ws.Cells(rowIndex, 13).Value
        statusValue = SafeText(ws.Cells(rowIndex, 14).Value)

        If statusValue <> "Closed" And IsDate(followUpDate) Then
            If CDate(followUpDate) < Date Then
                ws.Rows(rowIndex).Interior.Color = RGB(255, 199, 206)
            ElseIf CDate(followUpDate) = Date Then
                ws.Rows(rowIndex).Interior.Color = RGB(255, 235, 156)
            End If
        End If
    Next rowIndex
End Sub

Public Sub ShowTodayFollowUpsManual()
    ShowTodayFollowUps True
End Sub

Public Sub ShowTodayFollowUps(Optional ByVal forceMessage As Boolean = True)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim followUpDate As Variant
    Dim output As String

    Set ws = ThisWorkbook.Worksheets(SHEET_LEADS)
    lastRow = LastUsedRow(ws, 1)

    For rowIndex = 2 To lastRow
        followUpDate = ws.Cells(rowIndex, 13).Value
        If IsDate(followUpDate) Then
            If CDate(followUpDate) <= Date And SafeText(ws.Cells(rowIndex, 14).Value) <> "Closed" Then
                output = output & vbCrLf & ws.Cells(rowIndex, 1).Value & " | " & ws.Cells(rowIndex, 3).Value & " | " & ws.Cells(rowIndex, 10).Value
            End If
        End If
    Next rowIndex

    If Len(output) > 0 Then
        MsgBox "Due follow-ups:" & vbCrLf & output, vbInformation, APP_TITLE
    ElseIf forceMessage Then
        MsgBox "No follow-ups are due today.", vbInformation, APP_TITLE
    End If
End Sub

Private Function DefaultOwner() As String
    DefaultOwner = SafeText(ThisWorkbook.Worksheets(SHEET_SETTINGS).Range("B2").Value)
    If Len(DefaultOwner) = 0 Then
        DefaultOwner = Environ$("Username")
    End If
End Function
