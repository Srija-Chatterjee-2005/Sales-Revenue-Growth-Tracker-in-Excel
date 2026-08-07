Attribute VB_Name = "modValidation"
Option Explicit

Public Sub ValidateCurrentDataset()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim duplicateCount As Long
    Dim blankCount As Long
    Dim invalidRevenue As Long

    Set ws = ThisWorkbook.Worksheets("RAW_DATA")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "No imported data found.", vbExclamation, "SalesIQ Data Validation"
        Exit Sub
    End If

    blankCount = Application.WorksheetFunction.CountBlank(ws.Range("A2:A" & lastRow))
    duplicateCount = lastRow - 1 - Application.WorksheetFunction.SumProduct(1 / Application.WorksheetFunction.CountIf(ws.Range("A2:A" & lastRow), ws.Range("A2:A" & lastRow)))
    invalidRevenue = Application.WorksheetFunction.CountIf(ws.Range("L2:L" & lastRow), "<0")

    MsgBox "DATA QUALITY CHECK" & vbCrLf & vbCrLf & _
           "Records: " & Format$(lastRow - 1, "#,##0") & vbCrLf & _
           "Blank Order IDs: " & blankCount & vbCrLf & _
           "Estimated duplicate IDs: " & Application.Max(0, Round(duplicateCount, 0)) & vbCrLf & _
           "Negative Revenue Rows: " & invalidRevenue, _
           vbInformation, "SalesIQ Data Validation"
End Sub
