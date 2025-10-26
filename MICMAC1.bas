Attribute VB_Name = "MICMAC1"
Public ar() As Variant
Public Ind_power() As Double
Public Dep_power() As Double
Public names() As String

Sub SSIMExtract()

Dim ws As Worksheet, config As Worksheet
'Dim ar() As Variant, power() As Integer

Set ws = Worksheets("Structuring")
Set config = Worksheets("Configuration")


Dim tbl As ListObject
Dim row As Range

Set tbl = ws.ListObjects("SSIM")

Dim i As Integer, j As Integer
i = 1

'-1 because it starts in 0
ReDim names(tbl.ListRows.Count - 1)
ReDim ar(tbl.ListRows.Count - 1, tbl.ListRows.Count - 1)

For Each row In tbl.DataBodyRange.Rows 'for each row in the table ssim
    names(i - 1) = row.Cells(1, 1).Value 'save the variable name
    For j = i + 2 To tbl.ListColumns.Count 'check the upper diagonal values
       'save the info in numerical form in the matrix (struc = table without the names)
        If row.Cells(1, j).Value = "V" Then
            ar(i - 1, j - 2) = 1
            ar(j - 2, i - 1) = 0
        ElseIf row.Cells(1, j).Value = "A" Then
            ar(i - 1, j - 2) = 0
            ar(j - 2, i - 1) = 1
        ElseIf row.Cells(1, j).Value = "X" Then
            ar(i - 1, j - 2) = 1
            ar(j - 2, i - 1) = 1
        ElseIf row.Cells(1, j).Value = "O" Then
            ar(i - 1, j - 2) = 0
            ar(j - 2, i - 1) = 0
        Else
            MsgBox ("There are undefined relationships")
            Exit Sub
        End If
    Next j
    i = i + 1
Next row


Call Transitivity(ar)
Call Powers(ar, Ind_power, Dep_power)


'clear previous data
'config.Columns("F:H").ClearContents
'print variables names and power array
'tbl.DataBodyRange.Columns(1).Copy Destination:=config.range("F2")
'config.range("G2").Resize(UBound(power, 1) + 1, 2).Value = power

Call MICMAC(names, Ind_power, Dep_power)

End Sub



Sub Transitivity(ar() As Variant)

Dim i As Integer, j As Integer, Raux As Integer, Caux As Integer

f = UBound(ar, 1)
If UBound(ar, 1) = 1 Then
    Exit Sub
End If

For i = 0 To UBound(ar, 1) 'Loop through array rows
    For j = 0 To UBound(ar, 2) 'Loop through array columns
        If ar(i, j) = 0 And i <> j Then  'if there is no dependency and we are not looking at the variable connection with itself
            aux = 0
            'Will stop only if var. j is dependent on var. aux and aux is dependent on var. i OR if array ends
            Do While (ar(aux, j) <> 1 Or ar(i, aux) <> 1) And aux < UBound(ar, 1)
                aux = aux + 1
            Loop
            
            If ar(aux, j) = 1 And ar(i, aux) = 1 Then
                ar(i, j) = "1*"
            End If

        End If
    
    Next j
    
Next i



End Sub

Sub Powers(ar() As Variant, Ind_power() As Double, Dep_power() As Double)
'Dim powers(1, 1) As Integer
'will save the powers

ReDim Ind_power(UBound(ar, 1))
ReDim Dep_power(UBound(ar, 1))

For i = 0 To UBound(ar, 1)
    Ind_power(i) = 0
    For j = 0 To UBound(ar, 2) 'calculate the independent
            If i = 0 Then 'To only calculate once the dependent power
                Dep_power(j) = 0
                For row = 0 To UBound(ar, 1)
                    If ar(row, j) = 1 Or ar(row, j) = "1*" Then  'counts all 1's
                        Dep_power(j) = Dep_power(j) + 1   'dependence power (x value)
                    End If
                Next row
                Dep_power(j) = 10 * Dep_power(j) / UBound(ar, 1) 'rescale the values to a 0-10 scale
            End If
            If ar(i, j) = 1 Or ar(i, j) = "1*" Then
                Ind_power(i) = Ind_power(i) + 1 'independence power  (y value)
            End If
    Next j
    Ind_power(i) = 10 * Ind_power(i) / UBound(ar, 1) 'rescale the values to a 0-10 scale
Next i

  
    
End Sub
