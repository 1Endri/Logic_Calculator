VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Logic_Calculator 
   Caption         =   "Logic_Calculator"
   ClientHeight    =   8712.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6324
   OleObjectBlob   =   "Logic_Calculator.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Logic_Calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim logicExpression As String
Dim allowChange As Boolean

' Initialize the form with default settings
Private Sub Logic_Calculator_Initialize()
    allowChange = True ' Initialize to allow changes by default
End Sub

' Append text to the TextBox and update the logicExpression
Private Sub AppendText(textToAppend As String)
    If allowChange Then
        ' Check if there is a character in the TextBox
        If Len(TextBox1.text) > 0 Then
            Dim lastChar As String
            lastChar = Right(TextBox1.text, 1)

            ' Check if the last character is "T" or "F" and the new textToAppend is also "T" or "F"
            If (UCase(lastChar) = "T" Or UCase(lastChar) = "F") And (UCase(textToAppend) = "T" Or UCase(textToAppend) = "F") Then
                ' Do not append if both the last character and new textToAppend are "T" or "F"
                Exit Sub
            End If
        End If

        ' Append the text to the TextBox and logicExpression
        TextBox1.text = TextBox1.text & textToAppend
        logicExpression = logicExpression & textToAppend
    End If
End Sub

' Event handlers for buttons T, F, NOT, AND, OR, XOR, NAND, NOR, THEN, EQUIV, (, )
Private Sub CommandButtonT_Click()
    allowChange = True
    AppendText "T"
End Sub

Private Sub CommandButtonF_Click()
    allowChange = True
    AppendText "F"
End Sub

Private Sub CommandButtonNOT_Click()
    allowChange = True
    AppendText "NOT "
End Sub

Private Sub CommandButtonAND_Click()
    allowChange = True
    AppendText " AND "
End Sub

Private Sub CommandButtonOR_Click()
    allowChange = True
    AppendText " OR "
End Sub

Private Sub CommandButtonXOR_Click()
    allowChange = True
    AppendText " XOR "
End Sub

Private Sub CommandButtonNAND_Click()
    allowChange = True
    AppendText " NAND "
End Sub

Private Sub CommandButtonNOR_Click()
    allowChange = True
    AppendText " NOR "
End Sub

Private Sub CommandButtonTHEN_Click()
    allowChange = True
    AppendText " THEN "
End Sub

Private Sub CommandButtonEQUIV_Click()
    allowChange = True
    AppendText " EQUIV "
End Sub

Private Sub CommandButtonOpenBracket_Click()
    allowChange = True
    AppendText "("
End Sub

Private Sub CommandButtonCloseBracket_Click()
    allowChange = True
    AppendText ")"
End Sub

' Evaluate the expression and display the result or an error message
Private Sub CommandButtonEqual_Click()
    allowChange = True
    ' Check if the operator sequence is valid before evaluation
    If IsValidOperatorSequence(TextBox1.text) Then
        EvaluateExpression TextBox1.text
    Else
        MsgBox "Invalid expression", vbExclamation, "Error"
        TextBox1.text = ""
    End If
End Sub

' Clear the TextBox and logicExpression
Private Sub CommandButtonClear_Click()
    allowChange = True
    TextBox1.text = ""
    logicExpression = ""
End Sub

' Evaluate the expression and display the result or an error message
Private Sub EvaluateExpression(expression As String)
    On Error Resume Next
    Dim result As Variant
    result = EvaluateLogic(expression)
    On Error GoTo 0

    ' Check for correctness of the expression
    If IsExpressionCorrect(expression) Then
        If IsError(result) Then
            TextBox1.text = "Error: " & Err.Description
        Else
            TextBox1.text = result
            DisplayResultType EvaluateType(expression)
        End If
    Else
        MsgBox "The expression entered is wrong", vbExclamation, "Error"
        TextBox1.text = ""  ' Set TextBox1 to show nothing
    End If
End Sub

' Check if the expression is correct based on specific rules (e.g., balanced brackets)
Function IsExpressionCorrect(expression As String) As Boolean
    ' Implement your logic to check the correctness of the expression
    ' For example, you can check if there are balanced brackets, valid operators, etc.
    ' Modify this function based on your specific requirements.

    ' For simplicity, let's check if the number of open and close brackets is the same
    Dim openBracketCount As Integer
    Dim closeBracketCount As Integer
    Dim i As Integer ' Add this line to declare the variable i

    For i = 1 To Len(expression)
        If Mid(expression, i, 1) = "(" Then
            openBracketCount = openBracketCount + 1
        ElseIf Mid(expression, i, 1) = ")" Then
            closeBracketCount = closeBracketCount + 1
        End If
    Next i

    IsExpressionCorrect = (openBracketCount = closeBracketCount)
End Function

' Display the result type in a message box
Sub DisplayResultType(resultType As String)
    If resultType <> "" Then
        MsgBox resultType, vbInformation
    End If
End Sub

' Determine the result type based on specific patterns in the expression
Function EvaluateType(expression As String) As String
    Debug.Print "Expression: " & expression
    If InStr(1, expression, "T AND F", vbTextCompare) > 0 Then
        EvaluateType = "Contradiction"
    ElseIf InStr(1, expression, "T OR F", vbTextCompare) > 0 Then
        EvaluateType = "Tautology"
    ElseIf InStr(1, expression, "T XOR F", vbTextCompare) > 0 Then
        EvaluateType = "Exclusive Or"
    ElseIf InStr(1, expression, "NOT T", vbTextCompare) > 0 Then
        EvaluateType = "Negation"
    ElseIf InStr(1, expression, "NOT F", vbTextCompare) > 0 Then
        EvaluateType = "Negation"
    ElseIf InStr(1, expression, "T NAND F", vbTextCompare) > 0 Then
        EvaluateType = "Negation"
    ElseIf InStr(1, expression, "T NOR F", vbTextCompare) > 0 Then
        EvaluateType = "Negation"
    ElseIf InStr(1, expression, "T THEN F", vbTextCompare) > 0 Then
        EvaluateType = "Implication"
    ElseIf InStr(1, expression, "T EQUIV F", vbTextCompare) > 0 Then
        EvaluateType = "Biconditional"
    Else
        ' If the expression is not matched with any known patterns, don't display a message
        Exit Function
    End If
    Debug.Print "Result Type: " & EvaluateType
End Function

' Evaluate the logic expression and return the result
Function EvaluateLogic(expression As String) As Variant
    Dim parts() As String
    parts = Split(expression, " ")

    Dim i As Integer
    Dim result As Boolean
    Dim operator As String
    Dim firstOperandSet As Boolean

    ' Initialize result based on the first operand
    If UBound(parts) >= 0 Then
        If UCase(parts(0)) = "T" Then
            result = True
            firstOperandSet = True
        ElseIf UCase(parts(0)) = "F" Then
            result = False
            firstOperandSet = True
        End If
    End If

    ' Evaluate the logic expression
    For i = LBound(parts) To UBound(parts)
        If IsOperator(parts(i)) Then
            operator = parts(i)
        Else
            Dim operand As Boolean
            operand = IIf(UCase(parts(i)) = "T", True, False)

            If Not firstOperandSet Then
                firstOperandSet = True
                result = operand
            Else
                Select Case operator
                    Case "AND"
                        result = result And operand
                    Case "OR"
                        result = result Or operand
                    Case "XOR"
                        result = (result Xor operand)
                    Case "NOT"
                        result = Not operand
                    Case "NAND"
                        result = Not (result And operand)
                    Case "NOR"
                        result = Not (result Or operand)
                    Case "THEN"
                        result = Not result Or operand
                    Case "EQUIV"
                        result = (result And operand) Or (Not result And Not operand)
                End Select
            End If
        End If
    Next i

    ' Return the result of the evaluation
    EvaluateLogic = result
End Function

' Check if the operator sequence is valid
Function IsValidOperatorSequence(expression As String) As Boolean
    Dim parts() As String
    parts = Split(expression, " ")

    Dim i As Integer
    Dim lastPart As String
    Dim currentPart As String

    ' Iterate through the parts to check the validity of the operator sequence
    For i = LBound(parts) To UBound(parts)
        ' Skip the first iteration
        If i = LBound(parts) Then
            GoTo NextIteration
        End If

        ' Get the last and current parts
        lastPart = parts(i - 1)
        currentPart = parts(i)

        ' Check if the current part is an operator
        If IsOperator(currentPart) Then
            ' If the last part is also an operator, the sequence is not valid
            If IsOperator(lastPart) Then
                IsValidOperatorSequence = False
                Exit Function
            End If

            ' If the last part is not an operator, it must be "T" or "F"
            If lastPart <> "T" And lastPart <> "F" Then
                IsValidOperatorSequence = False
                Exit Function
            End If
        End If
NextIteration:
    Next i

    ' If none of the checks failed, the sequence is valid
    IsValidOperatorSequence = True
End Function

' Check if a string represents an operator
Function IsOperator(s As String) As Boolean
    IsOperator = (s = "AND" Or s = "OR" Or s = "XOR" Or s = "NOT" Or s = "NAND" Or s = "NOR" Or s = "THEN" Or s = "EQUIV")
End Function

' Disable changes when TextBox is focused
Private Sub TextBox1_Enter()
    allowChange = False
End Sub

' Re-enable changes when TextBox loses focus
Private Sub TextBox1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    allowChange = True
End Sub

' Ignore all keyboard input by setting KeyAscii to 0
Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = 0
End Sub

' Click event for the UserForm
Private Sub UserForm_Click()
    ' You can add code here for the UserForm click event if needed
End Sub

