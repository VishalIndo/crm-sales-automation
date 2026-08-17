Attribute VB_Name = "modUtils"
Option Explicit

Public Function EnsureSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set EnsureSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If EnsureSheet Is Nothing Then
        Set EnsureSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        EnsureSheet.Name = sheetName
    End If
End Function

Public Function WorksheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    WorksheetExists = Not ws Is Nothing
End Function

Public Function LastUsedRow(ByVal ws As Worksheet, ByVal columnIndex As Long) As Long
    LastUsedRow = ws.Cells(ws.Rows.Count, columnIndex).End(xlUp).Row
End Function

Public Function EnsureTable(ByVal ws As Worksheet, ByVal tableName As String, ByVal headers As Variant) As ListObject
    Dim lastCol As Long
    Dim i As Long
    Dim tableRange As Range

    On Error Resume Next
    Set EnsureTable = ws.ListObjects(tableName)
    On Error GoTo 0

    If EnsureTable Is Nothing Then
        For i = LBound(headers) To UBound(headers)
            ws.Cells(1, i + 1).Value = headers(i)
            ws.Cells(1, i + 1).Font.Bold = True
        Next i

        lastCol = UBound(headers) - LBound(headers) + 1
        Set tableRange = ws.Range(ws.Cells(1, 1), ws.Cells(2, lastCol))
        Set EnsureTable = ws.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
        EnsureTable.Name = tableName
        EnsureTable.TableStyle = "TableStyleMedium2"
        EnsureTable.DataBodyRange.Rows(1).ClearContents
    End If
End Function

Public Function FindColumnIndex(ByVal tableObj As ListObject, ByVal headerName As String) As Long
    Dim i As Long

    For i = 1 To tableObj.ListColumns.Count
        If StrComp(tableObj.ListColumns(i).Name, headerName, vbTextCompare) = 0 Then
            FindColumnIndex = i
            Exit Function
        End If
    Next i

    Err.Raise vbObjectError + 700, "FindColumnIndex", "Header not found: " & headerName
End Function

Public Function NextId(ByVal prefix As String, ByVal ws As Worksheet, ByVal columnIndex As Long) As String
    Dim lastRow As Long
    Dim seq As Long
    Dim lastValue As String
    Dim parts() As String

    lastRow = LastUsedRow(ws, columnIndex)
    If lastRow < 2 Or Len(Trim$(ws.Cells(lastRow, columnIndex).Value)) = 0 Then
        seq = 1
    Else
        lastValue = CStr(ws.Cells(lastRow, columnIndex).Value)
        parts = Split(lastValue, "-")
        If UBound(parts) >= 2 And IsNumeric(parts(2)) Then
            seq = CLng(parts(2)) + 1
        Else
            seq = lastRow
        End If
    End If

    NextId = prefix & "-" & Format(Date, "yyyymmdd") & "-" & Format(seq, "000")
End Function

Public Function PromptText(ByVal title As String, ByVal message As String, Optional ByVal defaultValue As String = "") As String
    PromptText = Trim$(InputBox(message, title, defaultValue))
End Function

Public Function PromptDateValue(ByVal title As String, ByVal message As String, Optional ByVal defaultValue As Variant) As Variant
    Dim rawValue As String

    If IsMissing(defaultValue) Then
        rawValue = Trim$(InputBox(message, title, Format(Date, "yyyy-mm-dd")))
    Else
        rawValue = Trim$(InputBox(message, title, CStr(defaultValue)))
    End If

    If Len(rawValue) = 0 Then
        PromptDateValue = Empty
    ElseIf IsDate(rawValue) Then
        PromptDateValue = CDate(rawValue)
    Else
        MsgBox "Invalid date: " & rawValue, vbExclamation, APP_TITLE
        PromptDateValue = Empty
    End If
End Function

Public Function PromptNumberValue(ByVal title As String, ByVal message As String, Optional ByVal defaultValue As String = "") As Variant
    Dim rawValue As String

    rawValue = Trim$(InputBox(message, title, defaultValue))
    If Len(rawValue) = 0 Then
        PromptNumberValue = Empty
    ElseIf IsNumeric(rawValue) Then
        PromptNumberValue = CDbl(rawValue)
    Else
        MsgBox "Invalid number: " & rawValue, vbExclamation, APP_TITLE
        PromptNumberValue = Empty
    End If
End Function

Public Function SelectedLeadDataRow() As Long
    If ActiveSheet.Name <> SHEET_LEADS Then
        MsgBox "Select a lead row on the Leads sheet.", vbInformation, APP_TITLE
        Exit Function
    End If

    If ActiveCell.Row < 2 Then
        MsgBox "Select a data row, not the header.", vbInformation, APP_TITLE
        Exit Function
    End If

    SelectedLeadDataRow = ActiveCell.Row
End Function

Public Function JoinArray(ByVal values As Variant) As String
    Dim i As Long
    Dim result As String

    For i = LBound(values) To UBound(values)
        If Len(result) > 0 Then
            result = result & ","
        End If
        result = result & CStr(values(i))
    Next i

    JoinArray = result
End Function

Public Sub AutoFitUsedColumns(ByVal ws As Worksheet)
    ws.UsedRange.Columns.AutoFit
End Sub

Public Sub ApplyDateFormat(ByVal rng As Range)
    rng.NumberFormat = "yyyy-mm-dd"
End Sub

Public Sub ApplyMoneyFormat(ByVal rng As Range)
    rng.NumberFormat = "#,##0.00"
End Sub

Public Function GetLeadIdFromRow(ByVal rowNumber As Long) As String
    GetLeadIdFromRow = CStr(ThisWorkbook.Worksheets(SHEET_LEADS).Cells(rowNumber, 1).Value)
End Function

Public Function SafeText(ByVal value As Variant) As String
    If IsError(value) Then
        SafeText = ""
    ElseIf IsNull(value) Then
        SafeText = ""
    Else
        SafeText = Trim$(CStr(value))
    End If
End Function
