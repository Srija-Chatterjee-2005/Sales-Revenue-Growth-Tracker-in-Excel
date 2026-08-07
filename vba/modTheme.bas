Attribute VB_Name = "modTheme"
Option Explicit

Private Const DARK_BG As Long = 3349255
Private Const DARK_PANEL As Long = 472041
Private Const LIGHT_BG As Long = 16316664
Private Const LIGHT_PANEL As Long = 16777215

Public Sub ApplyDarkTheme()
    Dim ws As Worksheet

    Application.ScreenUpdating = False

    For Each ws In ThisWorkbook.Worksheets
        If IsUISheet(ws.Name) Then
            ws.Cells.Interior.Color = RGB(7, 27, 51)
            ws.Cells.Font.Color = RGB(255,255,255)
        End If
    Next ws

    RestyleDarkKeyAreas
    ApplyPremiumUI
    Application.ScreenUpdating = True

    MsgBox "Dark Premium theme applied.", vbInformation, "SalesIQ"
End Sub

Public Sub ApplyLightTheme()
    Dim ws As Worksheet

    Application.ScreenUpdating = False

    For Each ws In ThisWorkbook.Worksheets
        If IsUISheet(ws.Name) Then
            ws.Cells.Interior.Color = RGB(248,250,252)
            ws.Cells.Font.Color = RGB(15,23,42)
        End If
    Next ws

    RestyleLightKeyAreas
    ApplyPremiumUI
    Application.ScreenUpdating = True

    MsgBox "Light theme applied.", vbInformation, "SalesIQ"
End Sub

Private Function IsUISheet(ByVal sheetName As String) As Boolean
    Select Case sheetName
        Case "Home", "Dashboard", "KPI_Center", "Forecast", "Product_Analysis", _
             "Region_Analysis", "AI_Insights", "Reports", "Data_Dictionary", "Settings"
            IsUISheet = True
        Case Else
            IsUISheet = False
    End Select
End Function

Private Sub RestyleDarkKeyAreas()
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        If IsUISheet(ws.Name) Then
            On Error Resume Next
            ws.Range("A1:C50").Interior.Color = RGB(6,24,46)
            ws.Range("A1:C50").Font.Color = RGB(184,199,217)
            ws.Range("E2:P4").Interior.Color = RGB(7,27,51)
            ws.Range("E2:P4").Font.Color = RGB(255,255,255)
            On Error GoTo 0
        End If
    Next ws
End Sub

Private Sub RestyleLightKeyAreas()
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        If IsUISheet(ws.Name) Then
            On Error Resume Next
            ws.Range("A1:C50").Interior.Color = RGB(226,232,240)
            ws.Range("A1:C50").Font.Color = RGB(30,41,59)
            ws.Range("E2:P4").Interior.Color = RGB(241,245,249)
            ws.Range("E2:P4").Font.Color = RGB(15,23,42)
            On Error GoTo 0
        End If
    Next ws
End Sub
