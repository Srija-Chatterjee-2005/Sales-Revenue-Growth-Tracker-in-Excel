Attribute VB_Name = "modUI"
Option Explicit

Public Sub ApplyPremiumUI()
    Application.ScreenUpdating = False
    Application.DisplayFormulaBar = False
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.Zoom = 90
    Application.ScreenUpdating = True
End Sub

Public Sub ShowHome()
    ThisWorkbook.Worksheets("Home").Activate
    ApplyPremiumUI
End Sub

Public Sub ShowDashboard()
    ThisWorkbook.Worksheets("Dashboard").Activate
    ApplyPremiumUI
End Sub

Public Sub ShowKPI()
    ThisWorkbook.Worksheets("KPI_Center").Activate
    ApplyPremiumUI
End Sub

Public Sub ShowForecast()
    ThisWorkbook.Worksheets("Forecast").Activate
    ApplyPremiumUI
End Sub

Public Sub ShowReports()
    ThisWorkbook.Worksheets("Reports").Activate
    ApplyPremiumUI
End Sub

Public Sub ShowDictionary()
    ThisWorkbook.Worksheets("Data_Dictionary").Activate
    ApplyPremiumUI
End Sub

Public Sub ShowSettings()
    ThisWorkbook.Worksheets("Settings").Activate
    ApplyPremiumUI
End Sub
