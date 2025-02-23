Attribute VB_Name = "modFunctions"
Option Explicit

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
