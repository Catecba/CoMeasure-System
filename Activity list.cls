VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    

    Dim ws As Worksheet, config As Worksheet, tbl As ListObject
    Dim probChoice As String, probList As String, aux As Range
    Dim entries() As String
    Dim para As Integer
    Dim promptText As String
    Dim Parameters As Variant
    Dim Count As Variant
    Dim hasValidation As Boolean
    Set ws = Worksheets("Activity list")
    Set config = Worksheets("Configuration")
    Set tbl = ws.ListObjects("Activities")
    
    
    ' Safely check if cell has data validation. Avoids excel staying stuck on events = false

    On Error Resume Next
    hasValidation = (Target.Validation.Type <> 0)
    On Error GoTo ErrHandler

    
    'Check if the target is any of the cells bellow a CCR column
    If Not Application.Intersect(Target, tbl.DataBodyRange) Is Nothing And (Target.Count = 1) And Not IsEmpty(Target.Value) And hasValidation Then

        ' Check the probability distribution selected
        probChoice = Target.Value


        ' Determine dynamic prompt based on name
        Select Case LCase(probChoice) 'convert to lowercase
        
            Case "gamma"
                promptText = "Enter the Alpha and Beta parameters separated by semicolons"
                para = 2
                Parameters = InputBox(promptText, "Parameters Entry")
                ' Split input by semicolon
                entries = Split(Parameters, ";")
                
                
                '----- Check if empty or incorrect # entries
                If Parameters = "" Then
                    MsgBox "No input received", vbExclamation
                    Application.EnableEvents = False
                    Target.ClearContents
                    Target.Columns.AutoFit
                    Application.EnableEvents = True
                    Exit Sub
                ElseIf (UBound(entries) + 1) <> para Then  'InStr(Parameters, ",") > 0 Or InStr(Parameters, ".") > 0 Or
                      MsgBox "Invalied input. Each entry must be separated by a semicolon without empty values", vbExclamation
                      Application.EnableEvents = False
                      Target.ClearContents
                      Target.Columns.AutoFit
                      Application.EnableEvents = True
                      Exit Sub
                End If
                '-------
                
                text = "=GAMMA.INV(RAND();" & Trim(entries(0)) & ";" & Trim(entries(1)) & ")"
                Application.EnableEvents = False

                
                Target.FormulaLocal = text
                Application.EnableEvents = True
                
                
            Case "weibull"
                promptText = "Enter the Alpha and Beta parameters separated by semicolons"
                para = 2
                Parameters = InputBox(promptText, "Parameters Entry")
                ' Split input by semicolon
                entries = Split(Parameters, ";")
                
                '----- Check if only empty or incorrect # entries
                If Parameters = "" Then
                    MsgBox "No input received", vbExclamation
                    Application.EnableEvents = False
                    Target.ClearContents
                    Target.Columns.AutoFit
                    Application.EnableEvents = True
                    Exit Sub
                ElseIf (UBound(entries) + 1) <> para Then  'InStr(Parameters, ",") > 0 Or InStr(Parameters, ".") > 0 Or
                      MsgBox "Invalied input. Each entry must be separated by a semicolon without empty values", vbExclamation
                      Application.EnableEvents = False
                      Target.ClearContents
                      Target.Columns.AutoFit
                      Application.EnableEvents = True
                      Exit Sub
                End If
                '-------
                
                text = "=WEIBULL_INV(RAND();" & Trim(entries(0)) & ";" & Trim(entries(1)) & ")"
                Application.EnableEvents = False
                Target.FormulaLocal = text
                Application.EnableEvents = True
                
                
            Case "laplace"
                promptText = "Enter the Mean and Beta parameters separated by semicolons"
                para = 2
                Parameters = InputBox(promptText, "Parameters Entry")
                ' Split input by semicolon
                entries = Split(Parameters, ";")
                
                '----- Check if only empty or incorrect # entries
                If Parameters = "" Then
                    MsgBox "No input received", vbExclamation
                    Application.EnableEvents = False
                    Target.ClearContents
                    Target.Columns.AutoFit
                    Application.EnableEvents = True
                    Exit Sub
                ElseIf (UBound(entries) + 1) <> para Then  'InStr(Parameters, ",") > 0 Or InStr(Parameters, ".") > 0 Or
                      MsgBox "Invalied input. Each entry must be separated by a semicolon without empty values", vbExclamation
                      Application.EnableEvents = False
                      Target.ClearContents
                      Target.Columns.AutoFit
                      Application.EnableEvents = True
                      Exit Sub
                End If
                '-------
                
                text = "=LAPLACE_INV(RAND();" & Trim(entries(0)) & ";" & Trim(entries(1)) & ")"
                Application.EnableEvents = False
                Target.FormulaLocal = text
                Application.EnableEvents = True
                
                
            Case "lognorm"
                promptText = "Enter the Mean and Standard Deviation parameters separated by semicolons"
                para = 2
                Parameters = InputBox(promptText, "Parameters Entry")
                ' Split input by semicolon
                entries = Split(Parameters, ";")
                
                '----- Check if only empty or incorrect # entries
                If Parameters = "" Then
                    MsgBox "No input received", vbExclamation
                    Application.EnableEvents = False
                    Target.ClearContents
                    Target.Columns.AutoFit
                    Application.EnableEvents = True
                    Exit Sub
                ElseIf (UBound(entries) + 1) <> para Then  'InStr(Parameters, ",") > 0 Or InStr(Parameters, ".") > 0 Or
                      MsgBox "Invalied input. Each entry must be separated by a semicolon without empty values", vbExclamation
                      Application.EnableEvents = False
                      Target.ClearContents
                      Target.Columns.AutoFit
                      Application.EnableEvents = True
                      Exit Sub
                End If
                '-------
                
                text = "=LOGNORM.INV(RAND();" & Trim(entries(0)) & ";" & Trim(entries(1)) & ")"
                Application.EnableEvents = False
                Target.FormulaLocal = text
                Application.EnableEvents = True
                
                
            Case "normal"
                promptText = "Enter the Mean and Standard Deviation parameters separated by semicolons"
                para = 2
                Parameters = InputBox(promptText, "Parameters Entry")
                ' Split input by semicolon
                entries = Split(Parameters, ";")
                
                '----- Check if only empty or incorrect # entries
                If Parameters = "" Then
                    MsgBox "No input received", vbExclamation
                    Application.EnableEvents = False
                    Target.ClearContents
                    Target.Columns.AutoFit
                    Application.EnableEvents = True
                    Exit Sub
                ElseIf (UBound(entries) + 1) <> para Then  'InStr(Parameters, ",") > 0 Or InStr(Parameters, ".") > 0 Or
                      MsgBox "Invalied input. Each entry must be separated by a semicolon without empty values", vbExclamation
                      Application.EnableEvents = False
                      Target.ClearContents
                      Target.Columns.AutoFit
                      Application.EnableEvents = True
                      Exit Sub
                End If
                '-------
                
                text = "=NORM.INV(RAND();" & Trim(entries(0)) & ";" & Trim(entries(1)) & ")"
                Application.EnableEvents = False
                Target.FormulaLocal = text
                Application.EnableEvents = True
                
                
            Case "triang"
                promptText = "Enter the Min., Mode and Max. parameters separated by semicolons"
                para = 3
                Parameters = InputBox(promptText, "Parameters Entry")
                ' Split input by semicolon
                entries = Split(Parameters, ";")
                
                '----- Check if only empty or incorrect # entries
                If Parameters = "" Then
                    MsgBox "No input received", vbExclamation
                    Application.EnableEvents = False
                    Target.ClearContents
                    Target.Columns.AutoFit
                    Application.EnableEvents = True
                    Exit Sub
                ElseIf (UBound(entries) + 1) <> para Then  'InStr(Parameters, ",") > 0 Or InStr(Parameters, ".") > 0 Or
                      MsgBox "Invalied input. Each entry must be separated by a semicolon without empty values", vbExclamation
                      Application.EnableEvents = False
                      Target.ClearContents
                      Target.Columns.AutoFit
                      Application.EnableEvents = True
                      Exit Sub
                End If
                '-------
                
                text = "=TRIANG_INV(RAND();" & Trim(entries(0)) & ";" & Trim(entries(1)) & ";" & Trim(entries(2)) & ")"
                Application.EnableEvents = False
                Target.FormulaLocal = text
                Application.EnableEvents = True
                
            Case "pert"
                promptText = "Enter the Min., Mode and Max. parameters separated by semicolons"
                para = 3
                Parameters = InputBox(promptText, "Parameters Entry")
                ' Split input by semicolon
                entries = Split(Parameters, ";")
                
                '----- Check if only empty or incorrect # entries
                If Parameters = "" Then
                    MsgBox "No input received", vbExclamation
                    Application.EnableEvents = False
                    Target.ClearContents
                    Target.Columns.AutoFit
                    Application.EnableEvents = True
                    Exit Sub
                ElseIf (UBound(entries) + 1) <> para Then  'InStr(Parameters, ",") > 0 Or InStr(Parameters, ".") > 0 Or
                      MsgBox "Invalied input. Each entry must be separated by a semicolon without empty values", vbExclamation
                      Application.EnableEvents = False
                      Target.ClearContents
                      Target.Columns.AutoFit
                      Application.EnableEvents = True
                      Exit Sub
                End If
                '-------
                
                text = "=PERT_INV(RAND();" & Trim(entries(0)) & ";" & Trim(entries(1)) & ";" & Trim(entries(2)) & ")"
                Application.EnableEvents = False
                Target.FormulaLocal = text
                Application.EnableEvents = True
                
            Case "expon"
                promptText = "Enter the Lambda parameter"
                para = 1
                Parameters = InputBox(promptText, "Parameters Entry")
                ' Split input by semicolon
                entries = Split(Parameters, ";")
                
                '----- Check if empty or incorrect # entries
                If Parameters = "" Then
                    MsgBox "No input received", vbExclamation
                    Application.EnableEvents = False
                    Target.ClearContents
                    Target.Columns.AutoFit
                    Application.EnableEvents = True
                    Exit Sub
                ElseIf (UBound(entries) + 1) <> para Then  'InStr(Parameters, ",") > 0 Or InStr(Parameters, ".") > 0 Or
                      MsgBox "Invalied input. Each entry must be separated by a semicolon without empty values", vbExclamation
                      Application.EnableEvents = False
                      Target.ClearContents
                      Target.Columns.AutoFit
                      Application.EnableEvents = True
                      Exit Sub
                End If
                '-------
                
                text = "=EXPON_INV(RAND();" & Trim(entries(0)) & ")"
                Application.EnableEvents = False
                Target.FormulaLocal = text
                Application.EnableEvents = True
                
            Case "poisson"
                promptText = "Enter the Lambda parameter"
                para = 1
                Parameters = InputBox(promptText, "Parameters Entry")
                ' Split input by semicolon
                entries = Split(Parameters, ";")
                
                '----- Check if empty or incorrect # entries
                If Parameters = "" Then
                    MsgBox "No input received", vbExclamation
                    Application.EnableEvents = False
                    Target.ClearContents
                    Target.Columns.AutoFit
                    Application.EnableEvents = True
                    Exit Sub
                ElseIf (UBound(entries) + 1) <> para Then  'InStr(Parameters, ",") > 0 Or InStr(Parameters, ".") > 0 Or
                      MsgBox "Invalied input. Each entry must be separated by a semicolon without empty values", vbExclamation
                      Application.EnableEvents = False
                      Target.ClearContents
                      Target.Columns.AutoFit
                      Application.EnableEvents = True
                      Exit Sub
                End If
                '-------
                
                text = "=POISSON_INV(RAND();" & Trim(entries(0)) & ")"
                Application.EnableEvents = False
                Target.FormulaLocal = text
                Application.EnableEvents = True
                
            Case "uniform"
                promptText = "Enter the Min. and Max. parameters separated by semicolons"
                para = 2
                Parameters = InputBox(promptText, "Parameters Entry")
                ' Split input by semicolon
                entries = Split(Parameters, ";")
                
                '----- Check if empty or incorrect # entries
                If Parameters = "" Then
                    MsgBox "No input received", vbExclamation
                    Application.EnableEvents = False
                    Target.ClearContents
                    Target.Columns.AutoFit
                    Application.EnableEvents = True
                    Exit Sub
                ElseIf (UBound(entries) + 1) <> para Then  'InStr(Parameters, ",") > 0 Or InStr(Parameters, ".") > 0 Or
                      MsgBox "Invalied input. Each entry must be separated by a semicolon without empty values", vbExclamation
                      Application.EnableEvents = False
                      Target.ClearContents
                      Target.Columns.AutoFit
                      Application.EnableEvents = True
                      Exit Sub
                End If
                '-------
                
                text = "=UNIFORM_INV(RAND();" & Trim(entries(0)) & ";" & Trim(entries(1)) & ")"
                Application.EnableEvents = False
                Target.FormulaLocal = text
                Application.EnableEvents = True
                
            Case "binomial"
                promptText = "Enter the #trials and P(success) parameters separated by semicolons"
                para = 2
                Parameters = InputBox(promptText, "Parameters Entry")
                ' Split input by semicolon
                entries = Split(Parameters, ";")
            
                '----- Check if empty or incorrect # entries
                If Parameters = "" Then
                    MsgBox "No input received", vbExclamation
                    Application.EnableEvents = False
                    Target.ClearContents
                    Target.Columns.AutoFit
                    Application.EnableEvents = True
                    Exit Sub
                ElseIf (UBound(entries) + 1) <> para Then  'InStr(Parameters, ",") > 0 Or InStr(Parameters, ".") > 0 Or
                      MsgBox "Invalied input. Each entry must be separated by a semicolon without empty values", vbExclamation
                      Application.EnableEvents = False
                      Target.ClearContents
                      Target.Columns.AutoFit
                      Application.EnableEvents = True
                      Exit Sub
                End If
                '-------
                
                text = "=BINOM.INV(" & Trim(entries(0)) & ";" & Trim(entries(1)) & ";RAND())"
                Application.EnableEvents = False
                Target.FormulaLocal = text
                Application.EnableEvents = True
                
            Case "discrete"
                promptText1 = "Enter the Data Points parametersseparated by semicolons" & vbCrLf & "Example: ""10;20;30"""
                promptText2 = "Enter the Probabilities parameter separated by semicolons" & vbCrLf & "Example: ""0,2;0,5;0,3"""
                
                Parameters1 = InputBox(promptText1, "Data Points Parameter Entry")
                Parameters2 = InputBox(promptText2, "Probabilities Parameter Entry")
                
                
                '----- Check if empty or incorrect # entries
                If Parameters1 = "" Or Parameters2 = "" Then
                    MsgBox "No input received", vbExclamation
                    Application.EnableEvents = False
                    Target.ClearContents
                    Target.Columns.AutoFit
                    Application.EnableEvents = True
                    Exit Sub
                End If
                ' Split input by semicolon
                entries1 = Split(Parameters1, ";")
                entries2 = Split(Parameters2, ";")
                If UBound(entries1) <> UBound(entries2) Then 'if #data points <>#probabilities
                      MsgBox "Invalied input. Different number of Data Points and Probabilities", vbExclamation
                      Application.EnableEvents = False
                      Target.ClearContents
                      Target.Columns.AutoFit
                      Application.EnableEvents = True
                      Exit Sub
                End If
                '-------
                
                text = "=DISCR_INV(RAND();" & Parameters1 & ";" & Parameters2 & ")"
                Application.EnableEvents = False
                Target.FormulaLocal = text
                Application.EnableEvents = True
                
            Case "betageneral"
                promptText = "Enter the Alpha, Beta, Min. and Max. parameters separated by semicolons"
                para = 4
                Parameters = InputBox(promptText, "Parameters Entry")
                ' Split input by semicolon
                entries = Split(Parameters, ";")
                
                '----- Check if empty or incorrect # entries
                If Parameters = "" Then
                    MsgBox "No input received", vbExclamation
                    Application.EnableEvents = False
                    Target.ClearContents
                    Target.Columns.AutoFit
                    Application.EnableEvents = True
                    Exit Sub
                ElseIf (UBound(entries) + 1) <> para Then  'InStr(Parameters, ",") > 0 Or InStr(Parameters, ".") > 0 Or
                      MsgBox "Invalied input. Each entry must be separated by a semicolon without empty values", vbExclamation
                      Application.EnableEvents = False
                      Target.ClearContents
                      Target.Columns.AutoFit
                      Application.EnableEvents = True
                      Exit Sub
                End If
                '-------
                
                text = "=BETA.INV(RAND();" & Trim(entries(0)) & ";" & Trim(entries(1)) & ";" & Trim(entries(2)) & ";" & Trim(entries(3)) & ")"
                Application.EnableEvents = False
                Target.FormulaLocal = text
                Application.EnableEvents = True
                
            Case "singular-value"
                promptText = "Enter the time estimate for this component in the correct units"
                Parameters = InputBox(promptText, "Time Estimate Entry")
                Application.EnableEvents = False
                Target.ClearContents
                Target.Value = Parameters
                Application.EnableEvents = True
                Exit Sub
            

        End Select


    End If
    
Exit Sub

ErrHandler:
    Application.EnableEvents = True
    MsgBox "An error occurred: " & Err.Description, vbExclamation
    Target.ClearContents
End Sub




