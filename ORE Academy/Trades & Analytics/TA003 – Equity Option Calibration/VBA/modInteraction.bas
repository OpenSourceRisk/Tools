Attribute VB_Name = "modInteraction"
Option Explicit

Public Const defaultVolatility As Double = 0.15

'=========================== IRCURVES TAB ============================

'Increase the quote number by 1
Sub QuoteNumberUP()

    Dim quoteNumRange As Range: Set quoteNumRange = Range("rngCurrentQuoteNumber")
    Dim maxQuote As Integer: maxQuote = 37
    Dim newQuoteNum As Integer: newQuoteNum = Application.WorksheetFunction.Min(quoteNumRange.Value2 + 1, maxQuote)
    quoteNumRange.Value2 = newQuoteNum
    Range("rngRootFindingModifiableDF").Value2 = Range("rngRootFindingInitialGuessDF").Value2
    
End Sub

'Decrease the quote number by 1
Sub QuoteNumberDOWN()

    Dim quoteNumRange As Range: Set quoteNumRange = Range("rngCurrentQuoteNumber")
    Dim minQuote As Integer: minQuote = 1
    Dim newQuoteNum As Integer: newQuoteNum = Application.WorksheetFunction.Max(quoteNumRange.Value2 - 1, minQuote)
    quoteNumRange.Value2 = newQuoteNum
    Range("rngRootFindingModifiableDF").Value2 = Range("rngRootFindingInitialGuessDF").Value2
    
End Sub

'Reset Calibration to Initial Values
Sub resetCalibration()
    Range("rngCurrentQuoteNumber").Value2 = 1
    Range("rngRootFindingModifiableDF").Value2 = Range("rngRootFindingInitialGuessDF").Value2
End Sub

'=========================== EQIMPLIEDVOl TAB ============================

'Increase the Option quote number by 1 or by Option Type (i.e. CALL or PUT)
Sub EQQuoteNumberUP()

    Dim quoteNumRange As Range: Set quoteNumRange = Range("rngEQCurrentQuoteNumber")
    Dim quoteTypeRange As Range: Set quoteTypeRange = Range("rngEQCurrentQuoteType")
    Dim quoteNum As Integer: quoteNum = quoteNumRange.Value2
    Dim quoteType As String: quoteType = quoteTypeRange.Value2
    Dim maxQuote As Integer: maxQuote = 9
    Dim newQuoteNum As Integer
    Dim newQuoteType As String
    
    If quoteType = "CALL" Then
        newQuoteType = "PUT"
        quoteTypeRange.Value2 = newQuoteType
        Range("rngEQRootFindingModifiableDF").Value2 = defaultVolatility
    Else
        If quoteNum < maxQuote Then
            newQuoteNum = Application.WorksheetFunction.Min(quoteNum + 1, maxQuote)
            newQuoteType = "CALL"
            quoteNumRange.Value2 = newQuoteNum
            quoteTypeRange.Value2 = newQuoteType
            Range("rngEQRootFindingModifiableDF").Value2 = defaultVolatility
        End If
    End If
    
End Sub

'Decrease the Option quote number by 1 or by Option Type (i.e. CALL or PUT)
Sub EQQuoteNumberDOWN()

    Dim quoteNumRange As Range: Set quoteNumRange = Range("rngEQCurrentQuoteNumber")
    Dim quoteTypeRange As Range: Set quoteTypeRange = Range("rngEQCurrentQuoteType")
    Dim quoteNum As Integer: quoteNum = quoteNumRange.Value2
    Dim quoteType As String: quoteType = quoteTypeRange.Value2
    Dim minQuote As Integer: minQuote = 1
    Dim newQuoteNum As Integer
    Dim newQuoteType As String
    
    If quoteType = "PUT" Then
        newQuoteType = "CALL"
        quoteTypeRange.Value2 = newQuoteType
        Range("rngEQRootFindingModifiableDF").Value2 = defaultVolatility
    Else
        If quoteNum > minQuote Then
            newQuoteNum = Application.WorksheetFunction.Max(quoteNumRange.Value2 - 1, minQuote)
            newQuoteType = "PUT"
            quoteNumRange.Value2 = newQuoteNum
            quoteTypeRange.Value2 = newQuoteType
            Range("rngEQRootFindingModifiableDF").Value2 = defaultVolatility
        End If
    End If
    
End Sub

'Reset Option Calibration to Initial Values
Sub EQResetCalibration()
    Range("rngEQCurrentQuoteNumber").Value2 = 1
    Range("rngEQCurrentQuoteType").Value2 = "CALL"
    Range("rngEQRootFindingModifiableDF").Value2 = defaultVolatility
End Sub
