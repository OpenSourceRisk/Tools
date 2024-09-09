Attribute VB_Name = "modAddin"
Option Explicit

'Installation of the Excel Solver addin
Sub InstallExcelSolverAddin()
    
    'Search parameters
    Dim fileNameAddin As String: fileNameAddin = "Solver.xla"
    Dim excelNameAddin As String: excelNameAddin = "Solver Add-in"
    Dim excelReferenceNameAddin As String: excelReferenceNameAddin = "SOLVER"
    
    'Install the addin in Excel and VBA
    Call installAddin(fileNameAddin, excelNameAddin, excelReferenceNameAddin)
    
End Sub

'Installation of any Excel addin
'Make sure you have activated "Trust Access":
'In Excel, go to Files then Option/Trust Center click on Trust Center Settings then Macro Settings and tick Enable VBA macros
Sub installAddin(fileNameAddin As String, _
                 excelNameAddin As String, _
                 excelReferenceNameAddin As String)
    
    Dim objfilesearch As Variant
    Dim i As Integer
    Dim pathToExcelApplication As String: pathToExcelApplication = Application.Path
    Dim loadingAddinCommand As String: loadingAddinCommand = fileNameAddin & "!auto_open"
    
    'Create error management: if error is triggered then dismiss it and goes to next line
    On Error Resume Next
    
    Set objfilesearch = Application.FileSearch
    
    'Search for the add-in on the computer
    With objfilesearch
        .NewSearch
        .LookIn = pathToExcelApplication
        .SearchSubFolders = True
        .Filename = fileNameAddin
        .Execute
        
        'Case where the add-in WAS found on the computer
        If .Execute > 0 Then
        
            'Case where the add-in has not been installed previously
            If AddIns(excelNameAddin).Installed = False Then
                
                'Error management when the add-in could NOT be installed
                If Err.Number > 0 Then
                
                    'Case 1 - Due to Excel security settings
                    If Err.Number = 1004 Then
                        MsgBox "The Excel " & excelNameAddin & " add-in could not be installed due to Security Settings. Make sure you allow Macros in Trust Center Settings"
                        Err.Clear
                        Exit Sub
                        
                    'Case 2 - Due to the addin not listed in the add-ins list
                    Else
                        AddIns.Add(.FoundFiles(1)).Installed = True
                        Err.Clear
                    End If
                    
                End If
                
                'When no error was found during installation
                AddIns(excelNameAddin).Installed = True
                MsgBox "The Excel " & excelNameAddin & " add-in was successfully installed"
                
            End If
            
        'Case where the addin WAS NOT found on the computer
        Else
            MsgBox "The Excel " & excelNameAddin & " add-in could not be found on this computer", vbCritical
            Exit Sub
        End If
        
    End With
    
    'Check if the Excel addin has also been installed in the VBA
    For i = 1 To ThisWorkbook.VBProject.References.Count
        If ThisWorkbook.VBProject.References(i).name = excelReferenceNameAddin Then Exit Sub
    Next i
    
    'In case it has NOT been added yet, we install it in VBA
    ThisWorkbook.VBProject.References.AddFromFile Application.LibraryPath & _
                                                  Application.PathSeparator & _
                                                  excelReferenceNameAddin & _
                                                  Application.PathSeparator & UCase(fileNameAddin)
    
    'Finally, load the add-in
    Application.Run loadingAddinCommand
    
    'Remove error management
    On Error GoTo 0
    
End Sub


