Attribute VB_Name = "InvFunction"
Function DISCR_INV(p As Double, valueStr As String, probStr As String) As Variant
    Dim values() As String, probs() As String
    Dim i As Long, cumulative As Double
    Dim val As Double, prob As Double

    ' Split the input strings by semicolon
    values = Split(valueStr, ";")
    probs = Split(probStr, ";")
     
    'if the arrays have different sizes, error
    If UBound(values) <> UBound(probs) Then
        DISCR_INV = CVErr(xlErrValue)
        Exit Function
    End If
    'if probability not between 0 and 1, error
    If p < 0 Or p > 1 Then
        DISCR_INV = CVErr(xlErrNum)
        Exit Function
    End If

    cumulative = 0
    For i = 0 To UBound(values)
        val = values(i)
        prob = probs(i)
        
        'if value prob is not a number, error
        If Not IsNumeric(prob) Then
            DISCR_INV = CVErr(xlErrValue)
            Exit Function
        End If
        
        cumulative = cumulative + prob
        'if the prob is in the value prob interval: cum< p <= cum+ val prob
        If p <= cumulative Then
            DISCR_INV = val
            Exit Function
        End If
    Next i
    
    DISCR_INV = values(UBound(values)) ' returns last value if previous loop has ended and no value was selected
End Function


