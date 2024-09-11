Attribute VB_Name = "modCalibration"
Option Explicit

'Perform an Excel Goal Seek root-finding algorithm to estimate the implied volatility of each Call option
Sub getImpliedVolForCallOptions()
    
    Dim refHeaderCell As Range: Set refHeaderCell = Range("rngEQImpliedVolCalibAnchor")
    Dim countQuotes As Integer: countQuotes = Range("rngNbEqVolQuotes").Value2
    
    Call executeGoalSeek(refHeaderCell, countQuotes, "Call Price", "Calculated Call Price", "Call Volatility")
    
End Sub

'Perform an Excel Goal Seek root-finding algorithm to estimate the implied volatility of each Put option
Sub getImpliedVolForPutOptions()
    
    Dim refHeaderCell As Range: Set refHeaderCell = Range("rngEQImpliedVolCalibAnchor")
    Dim countQuotes As Integer: countQuotes = Range("rngNbEqVolQuotes").Value2
    
    Call executeGoalSeek(refHeaderCell, countQuotes, "Put Price", "Calculated Put Price", "Put Volatility")
    
End Sub

'Set the Solver Parameters (objective, engine type, number of steps...)
Sub setSolverParameters(nameRootFindingTargetCell As String, _
                        nameRootFindingModifiableCell As String, _
                        nameRootFindingAccuracyCell As String)
    
    Dim precisionSolver As Double: precisionSolver = Range(nameRootFindingAccuracyCell).Value2
    
    SolverOptions _
        MaxTime:=0, _
        Iterations:=60, _
        Precision:=precisionSolver, _
        Convergence:=precisionSolver, _
        StepThru:=True, _
        Scaling:=False, _
        AssumeNonNeg:=True, _
        Derivatives:=2, _
        PopulationSize:=100, _
        RandomSeed:=0, _
        MutationRate:=0.075, _
        Multistart:=False, _
        RequireBounds:=False, _
        MaxSubproblems:=0, _
        MaxIntegerSols:=0, _
        IntTolerance:=1, _
        SolveWithout:=False, _
        MaxTimeNoImp:=30
    
    SolverOk SetCell:=nameRootFindingTargetCell, _
             MaxMinVal:=3, _
             ValueOf:=0, _
             ByChange:=nameRootFindingModifiableCell, _
             Engine:=1, _
             EngineDesc:="GRG Nonlinear"""
        
End Sub

'Launch the animation performing the calibration of all 37 quotes, one after the other
Sub RootFindingAllIRQuote()
    
    Dim i As Integer
    Dim useSolver As Boolean: useSolver = Range("rngUseSolver").Value2
    Dim nameRootFindingTargetCell As String
    Dim nameRootFindingTargetCellValue As String: nameRootFindingTargetCellValue = ""
    Dim nameRootFindingModifiableCell As String: nameRootFindingModifiableCell = "rngRootFindingModifiableDF"
    Dim nameRootFindingAccuracyCell As String: nameRootFindingAccuracyCell = "rngArbitrageThresholdYieldCurve"
    Dim protectSheets As Integer: protectSheets = Range("rngProtectWorksheets").Value2
    Dim nameWorksheet As String: nameWorksheet = "IRCurve"
    
    If protectSheets = 1 Then Call unprotectSingleSheets(nameWorksheet)
    
    '1. Parametrize the root-finding/solver algorithm
    If useSolver Then
        nameRootFindingTargetCellValue = "rngRootFindingTargetValue"
        Call setSolverParameters(nameRootFindingTargetCellValue, nameRootFindingModifiableCell, nameRootFindingAccuracyCell)
    Else
        
        If Range("rngOptimisationTargetType").Value2 = "Rate Quote" Then
            nameRootFindingTargetCell = "rngEstimatedRate"
            nameRootFindingTargetCellValue = "rngTargetRate"
        Else
            nameRootFindingTargetCell = "rngFloatingLegNPV2"
            nameRootFindingTargetCellValue = "rngFixLegNPV"
        End If
    End If
                        
    '1. Parametrize the root-finding/solver algorithm
    If useSolver Then Call setSolverParameters(nameRootFindingTargetCell, nameRootFindingModifiableCell, nameRootFindingAccuracyCell)
    
    '2. Reset to first quote, refresh and update graph
    Call resetCalibration
    DoEvents
    Application.CalculateFullRebuild
    DoEvents
    Application.Wait (Now + TimeValue("0:00:01") / 2)
    Call ChangeCharts
    DoEvents
    
    '3. Go through each quote chronologically and run the root-finding algorithm for that quote
    For i = 1 To 37
        
        '4. Go to next quote
        If i > 1 Then
            Call QuoteNumberUP
            DoEvents
        End If
        
        '5. Run the root-finding algorithm for that quote
        Call RootFindingIndividualIRQuote(nameRootFindingTargetCell, nameRootFindingModifiableCell, nameRootFindingTargetCellValue, useSolver)
        
        '6. Save the implied discount factor calculated
        Range("rngESTERSavingTableHeader").Offset(i + 1, 0).Value2 = Range(nameRootFindingModifiableCell).Value2
        
        '7. Refresh the spreadsheet
        Application.CalculateFullRebuild
        DoEvents
        Application.Wait (Now + TimeValue("0:00:01") / 2)
        
    Next i
    
    If protectSheets = 1 Then Call protectSingleSheets(nameWorksheet)
    
End Sub


'Run the root-finding/solver algorithm for the current quote
Sub runRootFindingFindingIndividualIRQuote()
    
    Dim useSolver As Boolean: useSolver = Range("rngUseSolver").Value2
    Dim nameRootFindingTargetCell As String
    Dim nameRootFindingTargetCellValue As String: nameRootFindingTargetCellValue = ""
    Dim nameRootFindingModifiableCell As String: nameRootFindingModifiableCell = "rngRootFindingModifiableDF"
    Dim nameRootFindingAccuracyCell As String: nameRootFindingAccuracyCell = "rngArbitrageThresholdYieldCurve"
    Dim protectSheets As Integer: protectSheets = Range("rngProtectWorksheets").Value2
    Dim nameWorksheet As String: nameWorksheet = "IRCurve"
    
    If protectSheets = 1 Then Call unprotectSingleSheets(nameWorksheet)
    
    '1. Parametrize the root-finding/solver algorithm
    If useSolver Then
        nameRootFindingTargetCellValue = "rngRootFindingTargetValue"
        Call setSolverParameters(nameRootFindingTargetCellValue, nameRootFindingModifiableCell, nameRootFindingAccuracyCell)
    Else
        
        If Range("rngOptimisationTargetType").Value2 = "Rate Quote" Then
            nameRootFindingTargetCell = "rngEstimatedRate"
            nameRootFindingTargetCellValue = "rngTargetRate"
        Else
            nameRootFindingTargetCell = "rngFloatingLegNPV2"
            nameRootFindingTargetCellValue = "rngFixLegNPV"
        End If
    End If
    
    '2. Run the solver
    Call RootFindingIndividualIRQuote(nameRootFindingTargetCell, nameRootFindingModifiableCell, nameRootFindingTargetCellValue, useSolver)
    
    If protectSheets = 1 Then Call protectSingleSheets(nameWorksheet)
    
End Sub

'Run the root-finding/solver algorithm for the current quote
Sub RootFindingIndividualIRQuote(nameRootFindingTargetCell As String, _
                                 nameRootFindingModifiableCell As String, _
                                 nameRootFindingTargetCellValue As String, _
                                 Optional useSolver As Boolean = False)
    
    Dim targetValue As Double: targetValue = Range(nameRootFindingTargetCellValue).Value2
    
    'Run the root-finding algorithm
    If useSolver Then
        SolverSolve UserFinish:=True, ShowRef:="SolverDisplayFunction"
    Else
        
        Call executeGoalSeekBootstrapping(Range(nameRootFindingTargetCell), targetValue, Range(nameRootFindingModifiableCell))
    End If
    
    'Refresh the spreadsheet
    Application.CalculateFullRebuild
    DoEvents
    
    'Update the charts
    Call ChangeCharts
    
End Sub

'Alpha function triggered during the root-finding/solver algorithm
Function SolverDisplayFunction(Reason As Integer)

     ThisWorkbook.RefreshAll
     DoEvents
     'Application.Wait (Now + TimeValue("0:00:02"))
     SolverDisplayFunction = 0
     
End Function

' This function's purpose is to adjust the charts' axess scale so that we zoomed on the data in an optimal way
Sub ChangeCharts()
    
    'The code below is just related to visuals so we can skip if an error occures
    On Error Resume Next
    
    Dim irCurveSheet As Worksheet: Set irCurveSheet = Worksheets("IRCurve")
    Dim scaleAxesRatesMinMax As Double: scaleAxesRatesMinMax = 0.075
    Dim scaleAxesDFMinMax As Double: scaleAxesDFMinMax = 0.00075
    
    irCurveSheet.Activate
    
    'Refresh all chart in the IRCurve tab
    DoEvents
    Application.ScreenUpdating = False
    Dim myChart As ChartObject
    For Each myChart In irCurveSheet.ChartObjects
        myChart.Chart.Refresh
    Next myChart
    
    'These DoEvents are sporadically placed to make sure the chart updates correclty as Excel tends to have issues with this
    DoEvents
    
    irCurveSheet.ChartObjects("Chart 3").Activate
    ActiveChart.FullSeriesCollection(2).DataLabels.Select
    Selection.AutoText = True
    ActiveChart.FullSeriesCollection(3).DataLabels.Select
    Selection.AutoText = True
    DoEvents
    
    irCurveSheet.ChartObjects("Chart 8").Activate
    ActiveChart.Axes(xlValue).Select
    
    Dim minRateValue As Double: minRateValue = Range("rngGraphRateMin").Value2
    Dim maxRateValue As Double: maxRateValue = Range("rngGraphRateMax").Value2
    Dim scaledMinRateValue As Double
    Dim scaledMaxRateValue As Double
    
    If minRateValue < 0 Then scaledMinRateValue = minRateValue * (1 + scaleAxesRatesMinMax) Else scaledMinRateValue = minRateValue * (1 - scaleAxesRatesMinMax)
    If maxRateValue < 0 Then scaledMaxRateValue = maxRateValue * (1 - scaleAxesRatesMinMax) Else scaledMaxRateValue = maxRateValue * (1 + scaleAxesRatesMinMax)
    
    ActiveChart.Axes(xlValue).MinimumScale = scaledMinRateValue
    ActiveChart.Axes(xlValue).MaximumScale = scaledMaxRateValue
    
    irCurveSheet.ChartObjects("Chart 9").Activate
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MinimumScale = Range("rngGraphDFMin").Value2 * (1 - scaleAxesDFMinMax)
    ActiveChart.Axes(xlValue).MaximumScale = Range("rngGraphDFMax").Value2 * (1 + scaleAxesDFMinMax)
    
    Range("A1").Select
    
    Application.ScreenUpdating = True
    
    On Error GoTo 0
    
End Sub



