Attribute VB_Name = "modOptionPricing"
Option Explicit

'Measure of Moneyness calculation with Black-76 Model
Function d1CalcBlack76(EquityForward As Double, _
                       Strike As Double, _
                       TimeToMaturity As Double, _
                       Sigma As Double) As Double
    
    Dim sigmaT As Double: sigmaT = Sigma * Sqr(TimeToMaturity)
    
    d1CalcBlack76 = WorksheetFunction.Ln(EquityForward / Strike) / sigmaT + sigmaT / 2
    
End Function

'Measure of Moneyness calculation
Function d1calc(S As Double, _
            K As Double, _
            T As Double, _
            rfwd As Double, _
            Sigma As Double, _
            q As Double) As Double
    
    Dim F As Double: F = S * Exp(-q * T) / Exp(-rfwd * T)
    Dim sigmaT As Double: sigmaT = Sigma * Sqr(T)
    
    d1calc = WorksheetFunction.Ln(F / K) / sigmaT + sigmaT / 2
    
End Function

'Cumulative Distribution Function (NDF) of the normal distribution
Function NCalc(d) As Double
    
    NCalc = WorksheetFunction.Norm_Dist(d, 0, 1, True)

End Function

'Call Equity Option Pricing with Black-76
Function CallOptionBlack76(EquityForward As Double, _
                           Strike As Double, _
                           TimeToMaturity As Double, _
                           DF As Double, _
                           Sigma As Double)
    
    Dim d1 As Double: d1 = d1CalcBlack76(EquityForward, Strike, TimeToMaturity, Sigma)
    Dim d2 As Double: d2 = d1 - Sigma * Sqr(TimeToMaturity)
    Dim nd1 As Double: nd1 = NCalc(d1)
    Dim nd2 As Double: nd2 = NCalc(d2)
    
    CallOptionBlack76 = DF * (EquityForward * nd1 + Strike * -nd2)
    
End Function

'Put Equity Option Pricing with Black-76
Function PutOptionBlack76(EquityForward As Double, _
                           Strike As Double, _
                           TimeToMaturity As Double, _
                           DF As Double, _
                           Sigma As Double)
    
    Dim d1 As Double: d1 = d1CalcBlack76(EquityForward, Strike, TimeToMaturity, Sigma)
    Dim d2 As Double: d2 = d1 - Sigma * Sqr(TimeToMaturity)
    Dim nd1 As Double: nd1 = NCalc(-d1)
    Dim nd2 As Double: nd2 = NCalc(-d2)
    
    PutOptionBlack76 = DF * (Strike * nd2 - EquityForward * nd1)
    
End Function

'Call Equity Option Pricing
Function CallOption(S As Double, _
                    K As Double, _
                    T As Double, _
                    r As Double, _
                    rfwd As Double, _
                    Sigma As Double, _
                    q As Double)
    
    Dim d1 As Double: d1 = d1calc(S, K, T, rfwd, Sigma, q)
    Dim d2 As Double: d2 = d1 - Sigma * Sqr(T)
    Dim nd1 As Double: nd1 = NCalc(d1)
    Dim nd2 As Double: nd2 = NCalc(d2)
    Dim DF As Double: DF = Exp(-r * T)
    Dim dfq As Double: dfq = Exp(-q * T)
    Dim F As Double: F = S * dfq / Exp(-rfwd * T)
    
    CallOption = DF * (F * nd1 + K * -nd2)
    
End Function

'Put Equity Option Pricing
Function PutOption(S As Double, _
                    K As Double, _
                    T As Double, _
                    r As Double, _
                    rfwd As Double, _
                    Sigma As Double, _
                    q As Double)
    
    Dim d1 As Double: d1 = d1calc(S, K, T, rfwd, Sigma, q)
    Dim d2 As Double: d2 = d1 - Sigma * Sqr(T)
    Dim nd1 As Double: nd1 = NCalc(-d1)
    Dim nd2 As Double: nd2 = NCalc(-d2)
    Dim DF As Double: DF = Exp(-r * T)
    Dim dfq As Double: dfq = Exp(-q * T)
    Dim F As Double: F = S * dfq / Exp(-rfwd * T)
    
    PutOption = DF * (K * nd2 - F * nd1)
    
End Function

