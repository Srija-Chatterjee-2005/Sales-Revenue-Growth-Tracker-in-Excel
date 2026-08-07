Attribute VB_Name = "modButtons"
Option Explicit

Public Sub CreateSalesIQButtons()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Home")

    DeleteSalesIQButtons

    AddRoundedButton ws, "btnUploadCSV", "UPLOAD CSV", "UploadCSV", 305, 640, 145, 42, RGB(34, 211, 238)
    AddRoundedButton ws, "btnRefreshAll", "REFRESH ALL", "RefreshSalesIQ", 465, 640, 145, 42, RGB(41, 121, 255)
    AddRoundedButton ws, "btnExportPDF", "EXPORT PDF", "ExportDashboardPDF", 625, 640, 145, 42, RGB(139, 92, 246)
    AddRoundedButton ws, "btnClearData", "CLEAR DATA", "ClearRawData", 785, 640, 145, 42, RGB(239, 68, 68)

    AddRoundedButton ws, "btnDarkTheme", "DARK THEME", "ApplyDarkTheme", 305, 700, 305, 42, RGB(11, 39, 72)
    AddRoundedButton ws, "btnLightTheme", "LIGHT THEME", "ApplyLightTheme", 625, 700, 305, 42, RGB(226, 232, 240)

    AddNavigationButtons
End Sub

Private Sub AddNavigationButtons()
    Dim sheetNames As Variant
    Dim macroNames As Variant
    Dim captions As Variant
    Dim i As Long
    Dim ws As Worksheet

    sheetNames = Array("Home", "Dashboard", "KPI_Center", "Forecast", "Product_Analysis", "Region_Analysis", "AI_Insights", "Reports", "Data_Dictionary", "Settings")
    macroNames = Array("ShowHome", "ShowDashboard", "ShowKPI", "ShowForecast", "OpenProductAnalysis", "OpenRegionAnalysis", "OpenAIInsights", "ShowReports", "ShowDictionary", "ShowSettings")
    captions = Array("Home", "Dashboard", "KPI Center", "Forecast", "Product Analysis", "Region Analysis", "AI Insights", "Reports", "Data Dictionary", "Settings")

    For i = LBound(sheetNames) To UBound(sheetNames)
        Set ws = ThisWorkbook.Worksheets(CStr(sheetNames(i)))
        AddRoundedButton ws, "navHome", "HOME", "ShowHome", 8, 8, 72, 24, RGB(37, 99, 235)

        If CStr(sheetNames(i)) <> "Home" Then
            AddRoundedButton ws, "navCurrent", CStr(captions(i)), CStr(macroNames(i)), 86, 8, 130, 24, RGB(99, 102, 241)
        End If
    Next i
End Sub

Private Sub AddRoundedButton(ByVal ws As Worksheet, ByVal shapeName As String, ByVal caption As String, _
                             ByVal macroName As String, ByVal leftPos As Double, ByVal topPos As Double, _
                             ByVal buttonWidth As Double, ByVal buttonHeight As Double, ByVal fillColor As Long)
    Dim shp As Shape

    On Error Resume Next
    ws.Shapes(shapeName).Delete
    On Error GoTo 0

    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, buttonWidth, buttonHeight)

    With shp
        .Name = shapeName
        .Fill.ForeColor.RGB = fillColor
        .Line.ForeColor.RGB = fillColor
        .Shadow.Visible = msoTrue
        .Shadow.Blur = 6
        .Shadow.OffsetX = 1
        .Shadow.OffsetY = 2
        .OnAction = macroName

        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.Text = caption
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .TextRange.Font.Name = "Segoe UI"
            .TextRange.Font.Size = 11
            .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255,255,255)
        End With
    End With
End Sub

Public Sub DeleteSalesIQButtons()
    Dim ws As Worksheet
    Dim i As Long

    For Each ws In ThisWorkbook.Worksheets
        For i = ws.Shapes.Count To 1 Step -1
            If Left$(ws.Shapes(i).Name, 3) = "btn" Or Left$(ws.Shapes(i).Name, 3) = "nav" Then
                ws.Shapes(i).Delete
            End If
        Next i
    Next ws
End Sub
