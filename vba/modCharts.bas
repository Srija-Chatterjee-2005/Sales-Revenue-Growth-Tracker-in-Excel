Attribute VB_Name = "modCharts"
Option Explicit

Public Sub RefreshAllCharts()
    Dim ws As Worksheet
    Dim ch As ChartObject

    Application.ScreenUpdating = False
    Application.CalculateFull

    For Each ws In ThisWorkbook.Worksheets
        For Each ch In ws.ChartObjects
            On Error Resume Next
            ch.Chart.Refresh
            On Error GoTo 0
        Next ch
    Next ws

    Application.ScreenUpdating = True
End Sub

Public Sub OpenDashboard()
    ThisWorkbook.Worksheets("Dashboard").Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.Zoom = 90
End Sub

Public Sub OpenProductAnalysis()
    ThisWorkbook.Worksheets("Product_Analysis").Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
End Sub

Public Sub OpenRegionAnalysis()
    ThisWorkbook.Worksheets("Region_Analysis").Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
End Sub

Public Sub OpenAIInsights()
    ThisWorkbook.Worksheets("AI_Insights").Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
End Sub
