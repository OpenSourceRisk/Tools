Attribute VB_Name = "modProtection"
Option Explicit

'Dropdown reactive function: if set to true, all sheets are protected, if set to false, all sheets are unprotected
Sub DropDown12_Change()

    Dim protectSheets As Integer: protectSheets = Range("rngProtectWorksheets").Value2
    
    Application.ScreenUpdating = False
    If protectSheets = 1 Then
        Call protectAllSheets
    Else
        Call unprotectAllSheets
    End If
    Application.ScreenUpdating = True
    
End Sub

'Protect all sheets (without password)
Sub protectAllSheets()

    Dim ws As Worksheet
    Dim pwd As String
    
    For Each ws In Worksheets
        ws.Protect
    Next ws

End Sub

'Unprotect all sheets (without password)
Sub unprotectAllSheets()
    
    Dim currentWS As Worksheet: Set currentWS = ActiveSheet
    Dim ws As Worksheet
    Dim pwd As String
    
    For Each ws In Worksheets
        ws.Unprotect
    Next ws
    
    currentWS.Activate
    
End Sub

'Protect all sheets (without password)
Sub protectSingleSheets(sheetName As String)

    Dim pwd As String
    
    Worksheets(sheetName).Protect

End Sub

'Unprotect a single sheets (without password)
Sub unprotectSingleSheets(sheetName As String)

    Dim pwd As String

    Worksheets(sheetName).Unprotect

End Sub
