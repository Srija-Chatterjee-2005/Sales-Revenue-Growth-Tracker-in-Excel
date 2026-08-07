Attribute VB_Name = "modNavigation"
Option Explicit

Public Sub GoToHome()
    ThisWorkbook.Worksheets("Home").Activate
End Sub

Public Sub GoToDashboard()
    ThisWorkbook.Worksheets("Dashboard").Activate
End Sub

Public Sub GoToKPI()
    ThisWorkbook.Worksheets("KPI_Center").Activate
End Sub

Public Sub GoToForecast()
    ThisWorkbook.Worksheets("Forecast").Activate
End Sub

Public Sub GoToReports()
    ThisWorkbook.Worksheets("Reports").Activate
End Sub

Public Sub GoToDataDictionary()
    ThisWorkbook.Worksheets("Data_Dictionary").Activate
End Sub

Public Sub GoToSettings()
    ThisWorkbook.Worksheets("Settings").Activate
End Sub
