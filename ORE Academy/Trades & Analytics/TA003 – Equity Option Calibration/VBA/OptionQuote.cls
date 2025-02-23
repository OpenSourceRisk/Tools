VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OptionQuote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: OptionQuote

Private pExpiry As Date
Private pStrike As Double
Private pCallPut As String
Private pPrice As Double

' Expiry Property
Public Property Get expiry() As Date
    expiry = pExpiry
End Property

Public Property Let expiry(value As Date)
    pExpiry = value
End Property

' Strike Property
Public Property Get Strike() As Double
    Strike = pStrike
End Property

Public Property Let Strike(value As Double)
    pStrike = value
End Property

' CallPut Property (Character type)
Public Property Get CallPut() As String
    CallPut = pCallPut
End Property

Public Property Let CallPut(value As String)
    If Len(value) = 1 Then ' Ensure it's a single character
        pCallPut = value
    Else
        Err.Raise vbObjectError + 1, "OptionQuote", "CallPut must be a single character."
    End If
End Property

' Price Property
Public Property Get Price() As Double
    Price = pPrice
End Property

Public Property Let Price(value As Double)
    pPrice = value
End Property


