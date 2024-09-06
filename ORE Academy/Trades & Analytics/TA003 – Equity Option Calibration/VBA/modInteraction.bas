Attribute VB_Name = "modInteraction"
Option Explicit

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
