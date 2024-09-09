Attribute VB_Name = "modFiles"
Option Explicit

'Check if file exists. Can allow for a message to be returned if it doesnt
Function checkIfExist(filepath As String, _
                      Optional withMessage As Boolean = False) As Boolean
                      
    Dim fileFound As Boolean: fileFound = Dir(filepath) <> ""
    
    If withMessage And Not fileFound Then
        MsgBox ("Can't find file: " & filepath)
    End If
    
    checkIfExist = fileFound
    
End Function
