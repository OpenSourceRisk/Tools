Attribute VB_Name = "modStrikeSelection"
Option Explicit

'Find the best suited quote to calculate the equity forward price
Function estimateBestEquityForward(expirySelected As Date, _
                                   forecastDiscountFactor As Double, _
                                   quotesDataRange As Range)
    
    Dim quoteData As Variant: quoteData = quotesDataRange
    Dim callColl As New Collection
    Dim putColl As New Collection
    Dim strikeColl As New Collection
    Dim Strike As Double
    Dim quote As OptionQuote
    Dim i, j As Integer
    Dim insertionPosition As Long
    
    'STEP 1 - DATA COLLECTION
    'Create 2 collections containing the call and the put quotes data, only for the given expiry
    'This 2 collections will be ordered by strike value from smaller to higher
    'Please note: there should always be a call AND a put quote for each {expiry;strike} couple
    For i = 1 To UBound(quoteData, 1)
        Set quote = New OptionQuote
        quote.expiry = quoteData(i, 1)
        Strike = quoteData(i, 2)
        quote.Strike = Strike
        quote.CallPut = quoteData(i, 3)
        quote.Price = quoteData(i, 4)
        
        If quote.expiry = expirySelected Then
            'Loop through previous quote strikes and find the position where to insert it so that it is ordered
            insertionPosition = 0
            For j = 1 To callColl.Count
                If Strike < callColl.Item(j).Strike Then insertionPosition = j
            Next j
            
            If quote.CallPut = "C" Then
                If insertionPosition > 0 Then
                    callColl.Add Item:=quote, Before:=insertionPosition
                Else
                    callColl.Add Item:=quote
                End If
            Else
                If insertionPosition > 0 Then
                    putColl.Add Item:=quote, Before:=insertionPosition
                Else
                    putColl.Add Item:=quote
                End If
            End If
        End If
    Next i
    
    'STEP 2 - FORWARD GUESS ESTIMATION
    'We make a first guess at the forward price
    'Strikes are ordered, lowest to highest, we take the first guess as midpoint of 2 strikes
    'where (C-P) goes from positive to negative
    Dim forward As Double
    Dim nbStrikes As Long: nbStrikes = callColl.Count
    Dim callQuote As OptionQuote
    Dim putQuote As OptionQuote
    
    'If no quotes has C<=P, then the highest strike is kept
    forward = callColl.Item(nbStrikes).Strike
    For i = 1 To nbStrikes
        Set callQuote = callColl.Item(i)
        Set putQuote = putColl.Item(i)
        
        If callQuote.Price <= putQuote.Price Then
            If i = 1 Then 'First strike
                forward = callQuote.Strike
            Else 'Midpoint of this strike and previous strike
                forward = (callQuote.Strike + callColl.Item(i - 1).Strike) / 2
            End If
        End If
    Next i
    
    'STEP 3 - FORWARD CALCULATION & VERIFICATION
    Dim isFwd As Boolean: isFwd = False
    Dim maxIter As Integer: maxIter = 101
    Dim interpolationDataPrepared As Boolean: interpolationDataPrepared = False
    Dim newForward As Double
    
    'Data only needed in Case 3
    Dim arrayStrikes() As Variant
    Dim arrayCallPrices() As Variant
    Dim arrayPutPrices() As Variant
    Dim callPrice As Double
    Dim putPrice As Double
    
    j = 1
    While (Not (isFwd) And j < maxIter)
        'CASE 1 - If our guess is below the first strike we just take the relevant strike
        If forward <= callColl.Item(1).Strike Then
            Set callQuote = callColl.Item(1)
            Set putQuote = putColl.Item(1)
            'Calculate the forward using the callput forward parity
            newForward = forwardFromPutCallParity(callQuote.Strike, callQuote.Price, putQuote.Price, forecastDiscountFactor)
            'If forward is still less than first strike we accept this
            isFwd = newForward <= callQuote.Strike
        'CASE 2 - If our guess is after the last strike we just take the relevant strike
        ElseIf forward >= callColl.Item(nbStrikes).Strike Then
            Set callQuote = callColl.Item(nbStrikes)
            Set putQuote = putColl.Item(nbStrikes)
            'Calculate the forward using the callput forward parity
            newForward = forwardFromPutCallParity(callQuote.Strike, callQuote.Price, putQuote.Price, forecastDiscountFactor)
            'If forward is still greater than first strike we accept this
            isFwd = newForward >= callQuote.Strike
        'CASE 3 - If our guess is between the lowest and highest strike, we calculate a forward from interpolated call/put prices
        Else
            'We only built the interpolation data once
            If Not (interpolationDataPrepared) Then
                arrayStrikes = getArrayFromQuote(callColl, "Strike")
                arrayCallPrices = getArrayFromQuote(callColl, "Price")
                arrayPutPrices = getArrayFromQuote(putColl, "Price")
                interpolationDataPrepared = True
            End If
            'Interpolate Call & Put price for the given forward guess
            callPrice = InterpolateXArray(forward, arrayStrikes, arrayCallPrices, False, False)
            putPrice = InterpolateXArray(forward, arrayStrikes, arrayCallPrices, False, False)
            'Calculate the forward using the callput forward parity
            newForward = forwardFromPutCallParity(forward, callPrice, putPrice, forecastDiscountFactor)
            'Check: has it moved by less that 0.1%
            isFwd = Abs((newForward - forward) / forward) < 0.001
        End If
        forward = newForward
        j = j + 1
    Wend
    
    estimateBestEquityForward = forward
    
End Function
         
'Function calculating the equity forward using the call/put parity
Function forwardFromPutCallParity(Strike As Double, _
                                  callPrice As Double, _
                                  putPrice As Double, _
                                  forecastDiscountFactor As Double)
    
    forwardFromPutCallParity = Strike + (callPrice - putPrice) / forecastDiscountFactor
    
End Function

'Recover all the values associated to a given parameter from a collection of quote object
Function getArrayFromQuote(coll As Collection, _
                           fieldName As String)
             
    Dim quote As OptionQuote
    Dim i As Long: i = 1
    Dim outputArray() As Variant: ReDim outputArray(1 To coll.Count, 1 To 1)
    Dim valueToAdd As Variant
    
    For Each quote In coll
        Select Case fieldName
            Case "Expiry"
                valueToAdd = quote.expiry
            Case "Price"
                valueToAdd = quote.Price
            Case "Strike"
                valueToAdd = quote.Strike
            Case "CallPut"
                valueToAdd = quote.CallPut
        End Select
        outputArray(i, 1) = valueToAdd
        i = i + 1
    Next
             
    getArrayFromQuote = outputArray
         
End Function
                   
'Function checking if an key exists already in a collection
Function Exists(coll As Collection, _
                key As String) As Boolean

    On Error GoTo endFunction
    IsObject (coll.Item(key))
    Exists = True
    
endFunction:
End Function

'Function transforming a collection of string items to an 1D-array
Function collToArray(coll As Collection) As Variant()

    Dim nbItems As Long: nbItems = coll.Count
    Dim i As Long
    Dim outputArray() As Variant: ReDim outputArray(1 To nbItems)
    
    For i = 1 To nbItems
        outputArray(i) = coll.Item(i)
    Next i
    
    collToArray = outputArray
    
End Function

'Sorts a 1D-array from smallest to largest value using a standard bubble sort algorithm
Sub BubbleSort(inputArray() As Variant)

    Dim i As Long
    Dim j As Long
    Dim temp As Variant
     
    For i = 1 To UBound(inputArray) - 1
        For j = i + 1 To UBound(inputArray)
            If inputArray(i) > inputArray(j) Then
                temp = inputArray(j)
                inputArray(j) = inputArray(i)
                inputArray(i) = temp
            End If
        Next j
    Next i
    
End Sub

