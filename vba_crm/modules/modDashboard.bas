Attribute VB_Name = "modDashboard"
Option Explicit

Public Sub RefreshDashboard()
    Dim ws As Worksheet
    Dim leadsWs As Worksheet
    Dim quotesWs As Worksheet
    Dim lastLeadRow As Long
    Dim lastQuoteRow As Long
    Dim rowIndex As Long
    Dim stageCounts() As Long
    Dim stages As Variant
    Dim stageName As String
    Dim totalLeads As Long
    Dim openLeads As Long
    Dim dueToday As Long
    Dim wonLeads As Long
    Dim lostLeads As Long
    Dim openQuoteValue As Double
    Dim quotesThisMonth As Long
    Dim quoteDate As Variant
    Dim quoteStatus As String

    If Not WorksheetExists(SHEET_DASHBOARD) Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(SHEET_DASHBOARD)
    Set leadsWs = ThisWorkbook.Worksheets(SHEET_LEADS)
    Set quotesWs = ThisWorkbook.Worksheets(SHEET_QUOTES)
    lastLeadRow = LastUsedRow(leadsWs, 1)
    lastQuoteRow = LastUsedRow(quotesWs, 1)
    stages = StageList()
    ReDim stageCounts(LBound(stages) To UBound(stages))

    For rowIndex = 2 To lastLeadRow
        If Len(SafeText(leadsWs.Cells(rowIndex, 1).Value)) > 0 Then
            totalLeads = totalLeads + 1
            stageName = SafeText(leadsWs.Cells(rowIndex, 10).Value)

            If SafeText(leadsWs.Cells(rowIndex, 14).Value) = "Open" Then
                openLeads = openLeads + 1
            End If

            If IsDate(leadsWs.Cells(rowIndex, 13).Value) Then
                If CDate(leadsWs.Cells(rowIndex, 13).Value) <= Date And SafeText(leadsWs.Cells(rowIndex, 14).Value) <> "Closed" Then
                    dueToday = dueToday + 1
                End If
            End If

            If stageName = "Won" Then wonLeads = wonLeads + 1
            If stageName = "Lost" Then lostLeads = lostLeads + 1
            AccumulateStageCount stages, stageCounts, stageName
        End If
    Next rowIndex

    For rowIndex = 2 To lastQuoteRow
        quoteStatus = SafeText(quotesWs.Cells(rowIndex, 5).Value)
        If quoteStatus <> "Rejected" And quoteStatus <> "Expired" Then
            If IsNumeric(quotesWs.Cells(rowIndex, 4).Value) Then
                openQuoteValue = openQuoteValue + CDbl(quotesWs.Cells(rowIndex, 4).Value)
            End If
        End If

        quoteDate = quotesWs.Cells(rowIndex, 3).Value
        If IsDate(quoteDate) Then
            If Year(CDate(quoteDate)) = Year(Date) And Month(CDate(quoteDate)) = Month(Date) Then
                quotesThisMonth = quotesThisMonth + 1
            End If
        End If
    Next rowIndex

    ws.Range("B5").Value = totalLeads
    ws.Range("D5").Value = openLeads
    ws.Range("F5").Value = dueToday
    ws.Range("H5").Value = openQuoteValue
    ws.Range("H5").NumberFormat = "#,##0.00"

    For rowIndex = LBound(stages) To UBound(stages)
        ws.Cells(rowIndex + 10, 2).Value = stages(rowIndex)
        ws.Cells(rowIndex + 10, 3).Value = stageCounts(rowIndex)
    Next rowIndex

    ws.Range("E10").Value = wonLeads
    ws.Range("E11").Value = lostLeads
    ws.Range("C21").Value = quotesThisMonth
    ws.Range("C22").Value = openQuoteValue
    ws.Range("C22").NumberFormat = "#,##0.00"

    AutoFitUsedColumns ws
End Sub

Private Sub AccumulateStageCount(ByVal stages As Variant, ByRef counts() As Long, ByVal stageName As String)
    Dim i As Long

    For i = LBound(stages) To UBound(stages)
        If StrComp(CStr(stages(i)), stageName, vbTextCompare) = 0 Then
            counts(i) = counts(i) + 1
            Exit For
        End If
    Next i
End Sub
