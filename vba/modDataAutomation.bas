Attribute VB_Name = "modDataAutomation"
Option Explicit

Private Const RAW_SHEET As String = "RAW_DATA"
Private Const HOME_SHEET As String = "Home"
Private Const DASHBOARD_SHEET As String = "Dashboard"
Private Const RAW_TABLE As String = "tblSalesRaw"

Public Sub UploadCSV()
    Dim selectedFile As Variant
    Dim importBook As Workbook
    Dim importSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim rowsImported As Long
    Dim headers As Variant
    Dim i As Long

    On Error GoTo ImportError

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.StatusBar = "Waiting for CSV selection..."

    selectedFile = Application.GetOpenFilename( _
        FileFilter:="CSV Files (*.csv),*.csv", _
        Title:="SalesIQ - Select Sales CSV")

    If VarType(selectedFile) = vbBoolean Then GoTo SafeExit

    Set targetSheet = ThisWorkbook.Worksheets(RAW_SHEET)
    Set importBook = Workbooks.Open(CStr(selectedFile), Local:=True)
    Set importSheet = importBook.Worksheets(1)

    lastRow = importSheet.Cells(importSheet.Rows.Count, 1).End(xlUp).Row
    lastCol = importSheet.Cells(1, importSheet.Columns.Count).End(xlToLeft).Column

    If lastRow < 2 Then
        MsgBox "This CSV contains headers but no data rows.", vbExclamation, "SalesIQ"
        GoTo SafeExit
    End If

    headers = Array( _
        "Order ID", "Date", "Month", "Quarter", "Region", "City", _
        "Sales Executive", "Product Category", "Product Name", "Units Sold", _
        "Selling Price", "Revenue", "Cost", "Profit", "Profit Margin %", _
        "Customer Type", "Sales Channel", "Target Sales", "Achievement %", _
        "Growth %", "Order Status", "Discount %", "Rating")

    If lastCol <> UBound(headers) + 1 Then
        MsgBox "The CSV must contain exactly " & UBound(headers) + 1 & _
               " columns." & vbCrLf & _
               "Please use the sample CSV structure supplied with SalesIQ.", _
               vbCritical, "Invalid CSV Structure"
        GoTo SafeExit
    End If

    For i = LBound(headers) To UBound(headers)
        If Trim$(CStr(importSheet.Cells(1, i + 1).Value)) <> CStr(headers(i)) Then
            MsgBox "Header mismatch in column " & i + 1 & "." & vbCrLf & _
                   "Expected: " & headers(i) & vbCrLf & _
                   "Found: " & importSheet.Cells(1, i + 1).Value, _
                   vbCritical, "Invalid CSV Header"
            GoTo SafeExit
        End If
    Next i

    Application.StatusBar = "Validating data..."
    If Not ValidateImportedSheet(importSheet, lastRow) Then GoTo SafeExit

    Application.StatusBar = "Importing records..."
    ClearRawData False

    targetSheet.Range("A1").Resize(lastRow, lastCol).Value = _
        importSheet.Range("A1").Resize(lastRow, lastCol).Value

    rowsImported = lastRow - 1

    ResizeRawTable targetSheet, lastRow

    With targetSheet
        .Columns("B:B").NumberFormat = "yyyy-mm-dd"
        .Columns("K:N").NumberFormat = "₹#,##0"
        .Columns("R:R").NumberFormat = "₹#,##0"
        .Columns("O:O").NumberFormat = "0.0%"
        .Columns("S:T").NumberFormat = "0.0%"
        .Columns("V:V").NumberFormat = "0.0%"
        .Columns("W:W").NumberFormat = "0.0"
    End With

    Application.StatusBar = "Refreshing SalesIQ..."
    RefreshSalesIQ False

    ThisWorkbook.Worksheets(HOME_SHEET).Range("A39").Value = _
        "Last Updated" & vbCrLf & Format$(Now, "dd-mmm-yyyy") & vbCrLf & Format$(Now, "hh:mm AM/PM")

    MsgBox Format$(rowsImported, "#,##0") & _
           " sales records imported successfully." & vbCrLf & vbCrLf & _
           "SalesIQ has refreshed the dashboard, KPI Center, forecast and reports.", _
           vbInformation, "Import Complete"

    ThisWorkbook.Worksheets(DASHBOARD_SHEET).Activate

SafeExit:
    On Error Resume Next
    If Not importBook Is Nothing Then importBook.Close SaveChanges:=False
    Application.StatusBar = False
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ImportError:
    MsgBox "SalesIQ could not import the CSV." & vbCrLf & _
           Err.Description, vbCritical, "Import Error"
    Resume SafeExit
End Sub

Private Function ValidateImportedSheet(ByVal ws As Worksheet, ByVal lastRow As Long) As Boolean
    Dim r As Long
    Dim issueCount As Long

    issueCount = 0

    For r = 2 To lastRow
        If Len(Trim$(CStr(ws.Cells(r, 1).Value))) = 0 Then issueCount = issueCount + 1
        If Not IsDate(ws.Cells(r, 2).Value) Then issueCount = issueCount + 1
        If Not IsNumeric(ws.Cells(r, 10).Value) Then issueCount = issueCount + 1
        If Not IsNumeric(ws.Cells(r, 12).Value) Then issueCount = issueCount + 1
    Next r

    If issueCount > 0 Then
        If MsgBox(issueCount & " possible data-quality issues were detected." & vbCrLf & _
                  "Do you still want to import the file?", _
                  vbQuestion + vbYesNo, "Data Quality Warning") = vbNo Then
            ValidateImportedSheet = False
            Exit Function
        End If
    End If

    ValidateImportedSheet = True
End Function

Private Sub ResizeRawTable(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim lo As ListObject

    On Error Resume Next
    Set lo = ws.ListObjects(RAW_TABLE)
    On Error GoTo 0

    If lo Is Nothing Then
        Set lo = ws.ListObjects.Add( _
            SourceType:=xlSrcRange, _
            Source:=ws.Range("A1:W" & lastRow), _
            XlListObjectHasHeaders:=xlYes)
        lo.Name = RAW_TABLE
    Else
        lo.Resize ws.Range("A1:W" & lastRow)
    End If

    lo.TableStyle = "TableStyleMedium2"
End Sub

Public Sub RefreshSalesIQ(Optional ByVal showMessage As Boolean = True)
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim cn As WorkbookConnection

    On Error GoTo RefreshError

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.StatusBar = "Refreshing SalesIQ..."

    For Each cn In ThisWorkbook.Connections
        On Error Resume Next
        cn.Refresh
        On Error GoTo RefreshError
    Next cn

    ThisWorkbook.RefreshAll

    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.PivotCache.Refresh
            pt.RefreshTable
        Next pt
    Next ws

    Application.CalculateFullRebuild

    ThisWorkbook.Worksheets(HOME_SHEET).Range("A39").Value = _
        "Last Updated" & vbCrLf & Format$(Now, "dd-mmm-yyyy") & vbCrLf & Format$(Now, "hh:mm AM/PM")

    If showMessage Then
        MsgBox "Dashboard, KPI Center, Forecast and Reports refreshed.", _
               vbInformation, "SalesIQ Refresh Complete"
    End If

SafeExit:
    Application.StatusBar = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

RefreshError:
    MsgBox "Refresh failed: " & Err.Description, vbCritical, "SalesIQ"
    Resume SafeExit
End Sub

Public Sub ClearRawData(Optional ByVal showMessage As Boolean = True)
    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = ThisWorkbook.Worksheets(RAW_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If lastRow >= 2 Then ws.Range("A2:W" & lastRow).ClearContents

    On Error Resume Next
    ws.ListObjects(RAW_TABLE).Resize ws.Range("A1:W2")
    On Error GoTo 0

    If showMessage Then
        MsgBox "Imported records were removed. Column headers were retained.", _
               vbInformation, "SalesIQ"
    End If
End Sub

Public Sub ExportDashboardPDF()
    Dim outFile As String

    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "Save the workbook before exporting a PDF.", vbExclamation, "SalesIQ"
        Exit Sub
    End If

    outFile = ThisWorkbook.Path & Application.PathSeparator & _
              "SalesIQ_Dashboard_" & Format$(Now, "yyyymmdd_hhmm") & ".pdf"

    With ThisWorkbook.Worksheets(DASHBOARD_SHEET)
        .PageSetup.PrintArea = "$A$1:$Q$45"
        .PageSetup.Orientation = xlLandscape
        .PageSetup.Zoom = False
        .PageSetup.FitToPagesWide = 1
        .PageSetup.FitToPagesTall = 1
        .ExportAsFixedFormat _
            Type:=xlTypePDF, _
            Filename:=outFile, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=True
    End With

    MsgBox "Dashboard exported successfully.", vbInformation, "SalesIQ"
End Sub
