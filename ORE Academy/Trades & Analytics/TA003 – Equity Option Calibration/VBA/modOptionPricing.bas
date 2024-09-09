Attribute VB_Name = "modOptionPricing"
Option Explicit

'Measure of Moneyness calculation
Function d1calc(S As Double, _
            K As Double, _
            T As Double, _
            rfwd As Double, _
            sigma As Double, _
            q As Double) As Double
    
    Dim F As Double: F = S * Exp(-q * T) / Exp(-rfwd * T)
    Dim sigmaT As Double: sigmaT = sigma * Sqr(T)
    
    d1calc = WorksheetFunction.Ln(F / K) / sigmaT + sigmaT / 2
    
End Function

'Cumulative Distribution Function (NDF) of the normal distribution
Function NCalc(d) As Double
    
    NCalc = WorksheetFunction.Norm_Dist(d, 0, 1, True)

End Function

'Call Equity Option Pricing
Function CallOption(S As Double, _
                    K As Double, _
                    T As Double, _
                    r As Double, _
                    rfwd As Double, _
                    sigma As Double, _
                    q As Double)
    
    Dim d1 As Double: d1 = d1calc(S, K, T, rfwd, sigma, q)
    Dim d2 As Double: d2 = d1 - sigma * Sqr(T)
    Dim nd1 As Double: nd1 = NCalc(d1)
    Dim nd2 As Double: nd2 = NCalc(d2)
    Dim df As Double: df = Exp(-r * T)
    Dim dfq As Double: dfq = Exp(-q * T)
    Dim F As Double: F = S * dfq / Exp(-rfwd * T)
    
    CallOption = df * (F * nd1 + K * -nd2)
    
End Function

'Put Equity Option Pricing
Function PutOption(S As Double, _
                    K As Double, _
                    T As Double, _
                    r As Double, _
                    rfwd As Double, _
                    sigma As Double, _
                    q As Double)
    
    Dim d1 As Double: d1 = d1calc(S, K, T, rfwd, sigma, q)
    Dim d2 As Double: d2 = d1 - sigma * Sqr(T)
    Dim nd1 As Double: nd1 = NCalc(-d1)
    Dim nd2 As Double: nd2 = NCalc(-d2)
    Dim df As Double: df = Exp(-r * T)
    Dim dfq As Double: dfq = Exp(-q * T)
    Dim F As Double: F = S * dfq / Exp(-rfwd * T)
    
    PutOption = df * (K * nd2 - F * nd1)
    
End Function

