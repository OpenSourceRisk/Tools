Attribute VB_Name = "modInterpolation"
Option Explicit

'Perform a linear or log-linear interpolation
Function getLinearInterpolation(targetX As Variant, _
                                X1 As Variant, _
                                X2 As Variant, _
                                Y1 As Variant, _
                                Y2 As Variant, _
                                Optional isLogLinear As Boolean = False, _
                                Optional isLN As Boolean = True)
    Dim targetY As Double
    
    If isLogLinear Then
        If isLN Then
            targetY = WorksheetFunction.Ln(Y1) + (targetX - X1) * (WorksheetFunction.Ln(Y2) - WorksheetFunction.Ln(Y1)) / (X2 - X1)
            targetY = Exp(targetY)
        Else
            targetY = Application.Log10(Y1) + (targetX - X1) * (Application.Log10(Y2) - Application.Log10(Y1)) / (X2 - X1)
            targetY = 10 ^ targetY
        End If
    Else
        targetY = Y1 + (targetX - X1) * (Y2 - Y1) / (X2 - X1)
    End If
    
    getLinearInterpolation = targetY
    
End Function

'Linear or loglinear interpolation from range of values
Function InterpolateX(ByRef XInput As Range, _
                      ByRef XRangeInput As Range, _
                      ByRef YRangeInput As Range, _
                      Optional isLogLinear As Boolean = False) As Double
                      
    Dim X1, X2, PX1, PX2, Y1, Y2 As Variant
    Dim X As Variant: X = XInput.Value2
    Dim XRange() As Variant: XRange = XRangeInput
    Dim YRange() As Variant: YRange = YRangeInput
    
    XRange = convertRangeOfDatesToLong(XRange)
    YRange = convertRangeOfDatesToLong(YRange)
    
    Dim minX As Variant: minX = WorksheetFunction.Min(XRange)
    Dim maxX As Variant: maxX = WorksheetFunction.Max(XRange)
    
    'Convert to long if date
    If IsDate(X) Then X = Int(CDbl(X))
    
    'Case 1: X is smaller than minimum of X range
    If (X < minX) Then X = minX
    
    'Case 2: X is larger than maximum of X range
    If (X > maxX) Then X = maxX
    
        'Determine data table points around searchvalues
    X1 = WorksheetFunction.XLookup(X, XRange, XRange, , -1)
    X2 = WorksheetFunction.XLookup(X, XRange, XRange, , 1)
    
    PX1 = WorksheetFunction.Match(X1, XRange, 0)
    PX2 = WorksheetFunction.Match(X2, XRange, 0)
    
    If UBound(YRange, 1) > 1 Then
        Y1 = YRange(PX1, 1)
        Y2 = YRange(PX2, 1)
    Else
        Y1 = YRange(1, PX1)
        Y2 = YRange(1, PX2)
    End If
    
    If Y1 <> Y2 Then
        InterpolateX = getLinearInterpolation(X, X1, X2, Y1, Y2, isLogLinear)
    Else
        InterpolateX = Y1
    End If
    
End Function

'Bilinear interpolation
Function InterpolateXY(ByRef XInput As Range, _
                       YInput As Range, _
                       ByRef XRangeInput As Range, _
                       ByRef YRangeInput As Range, _
                       ByRef ValueTable As Range) As Double
        
    Dim X1, X2, Y1, Y2, PX1, PX2, PY1, PY2, V1, V2, V3, V4, FX, FY, C12, C34 As Variant
    Dim X As Variant: X = XInput.Value2
    Dim Y As Variant: Y = YInput.Value2
    Dim XRange() As Variant: XRange = XRangeInput
    Dim YRange() As Variant: YRange = YRangeInput
    
    XRange = convertRangeOfDatesToLong(XRange)
    YRange = convertRangeOfDatesToLong(YRange)
    
    Dim minX As Variant: minX = WorksheetFunction.Min(XRange)
    Dim maxX As Variant: maxX = WorksheetFunction.Max(XRange)
    Dim minY As Variant: minY = WorksheetFunction.Min(YRange)
    Dim maxY As Variant: maxY = WorksheetFunction.Max(YRange)
    
    'Convert to long if date
    If IsDate(X) Then X = Int(CDbl(X))
    If IsDate(Y) Then Y = Int(CDbl(Y))
    
    'Case 1: X is smaller than minimum of X range
    If (X < minX) Then X = minX
    
    'Case 2: X is larger than maximum of X range
    If (X > maxX) Then X = maxX
    
    'Case 1: Y is smaller than minimum of Y range
    If (Y < minY) Then Y = minY
    
    'Case 2: Y is larger than maximum of Y range
    If (Y > maxY) Then Y = maxY
    
    
    'Determine data table points around searchvalues
    X1 = WorksheetFunction.XLookup(X, XRange, XRange, , -1)
    X2 = WorksheetFunction.XLookup(X, XRange, XRange, , 1)
    Y1 = WorksheetFunction.XLookup(Y, YRange, YRange, , -1)
    Y2 = WorksheetFunction.XLookup(Y, YRange, YRange, , 1)
    
    
    PX1 = WorksheetFunction.Match(X1, XRange, 0)
    PX2 = WorksheetFunction.Match(X2, XRange, 0)
    PY1 = WorksheetFunction.Match(Y1, YRange, 0)
    PY2 = WorksheetFunction.Match(Y2, YRange, 0)
    
    V1 = ValueTable(PY1, PX1)
    V2 = ValueTable(PY1, PX2)
    V3 = ValueTable(PY2, PX1)
    V4 = ValueTable(PY2, PX2)
    
    'X and Y fractions
    If (X2 = X1) Then FX = 0 Else FX = (X - X1) / (X2 - X1)
    If (Y2 = Y1) Then FY = 0 Else FY = (Y - Y1) / (Y2 - Y1)

    '2 intermediate results after interpolation on X
    C12 = V1 + FX * (V2 - V1)
    C34 = V3 + FX * (V4 - V3)
    
    'final result after interpolation on Y
    InterpolateXY = C12 + FY * (C34 - C12)
    
End Function


