Attribute VB_Name = "modFunctions"
Option Explicit

'Execute a Newton-Raphson regression (i.e. the Excel GoalSeek feature)
Sub executeGoalSeek(refHeaderCell As Range, _
                    nbValuesToCalculate As Integer, _
                    nameCellRowToMatch As String, _
                    nameCellRowFormulaToChange As String, _
                    nameCellRowToChange As String)
    
    Dim cellToMatchRow As Integer: cellToMatchRow = WorksheetFunction.Match(nameCellRowToMatch, refHeaderCell.EntireColumn, 0)
    Dim cellToMatch As Range: Set cellToMatch = refHeaderCell.Offset(cellToMatchRow - refHeaderCell.Row, 2)
    Dim cellToFormulaChangeRow As Integer: cellToFormulaChangeRow = WorksheetFunction.Match(nameCellRowFormulaToChange, refHeaderCell.EntireColumn, 0)
    Dim cellToFormulaChange As Range: Set cellToFormulaChange = refHeaderCell.Offset(cellToFormulaChangeRow - refHeaderCell.Row, 2)
    Dim cellToChangeRow As Integer: cellToChangeRow = WorksheetFunction.Match(nameCellRowToChange, refHeaderCell.EntireColumn, 0)
    Dim cellToChange As Range: Set cellToChange = refHeaderCell.Offset(cellToChangeRow - refHeaderCell.Row, 2)
    Dim i As Integer
    
    Dim protectSheets As Integer: protectSheets = Range("rngProtectWorksheets").Value2
    Dim nameWorksheet As String: nameWorksheet = "EQImpliedVol"
    
    Application.ScreenUpdating = False
    
    If protectSheets = 1 Then Call unprotectSingleSheets(nameWorksheet)
    
    For i = 1 To nbValuesToCalculate
        Call executeGoalSeekBootstrapping(cellToFormulaChange, cellToMatch.Value2, cellToChange)
        Set cellToMatch = cellToMatch.Offset(0, 1)
        Set cellToFormulaChange = cellToFormulaChange.Offset(0, 1)
        Set cellToChange = cellToChange.Offset(0, 1)
    Next i
    
    If protectSheets = 1 Then Call protectSingleSheets(nameWorksheet)
    
    Application.ScreenUpdating = True

End Sub

'Execute a Newton-Raphson regression (i.e. the Excel GoalSeek feature)
Sub executeGoalSeekBootstrapping(cellToFormulaChange As Range, _
                                 valueToMatch As Variant, _
                                 cellToChange As Range)
                                 
    cellToFormulaChange.GoalSeek Goal:=valueToMatch, ChangingCell:=cellToChange
                    
End Sub



'If some values are dates, replace them by a long as there is issues with date in text
Function convertRangeOfDatesToLong(inputTable() As Variant)
    
    Dim outputTable() As Variant
    Dim nbRows As Integer:
    Dim nbCols As Integer:
    Dim i As Integer
    
    If IsDate(inputTable(1, 1)) Then
        nbRows = UBound(inputTable, 1)
        nbCols = UBound(inputTable, 2)
        ReDim outputTable(1 To nbRows, 1 To nbCols)
        If nbRows > nbCols Then
            For i = 1 To nbRows
                outputTable(i, 1) = Int(CDbl(inputTable(i, 1)))
            Next i
        Else
        For i = 1 To nbCols
                outputTable(1, i) = Int(CDbl(inputTable(1, i)))
            Next i
        End If
    Else
        outputTable = inputTable
    End If
    
    
    convertRangeOfDatesToLong = outputTable
    
End Function

'Logarithm Function
Function lnVBA(value As Double) As Double
    lnVBA = WorksheetFunction.Ln(value)
End Function

'Exponential Function
Function expVBA(value As Double) As Double
    expVBA = Exp(value)
End Function
