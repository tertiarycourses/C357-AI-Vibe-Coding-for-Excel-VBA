Attribute VB_Name = "Module2"
Option Explicit

' --- 1. SYNTAX ERROR: caught when you compile; the code will not run at all
'     This procedure is deliberately broken. It is COMMENTED OUT so the rest
'     of the project still compiles - uncomment it to reproduce the error.
'
' Sub SyntaxErrorDemo()
'     Dim x As String
'     x = UCase("hello"        ' <- missing closing parenthesis
'     MsgBox x
' End Sub


' --- 2. RUNTIME ERROR: valid code that fails while running (error 11)
Sub RuntimeErrorDemo()
    Dim a As Integer, b As Integer
    a = 100
    b = 0
    MsgBox a / b              ' <- division by zero stops execution here
End Sub


' --- 3. LOGICAL ERROR: runs perfectly and reports nothing,
'        but the answer is WRONG
Sub LogicalErrorDemo()
    Dim curCommission As Currency
    curCommission = ActiveCell.Value * 0.01    ' <- should be 0.1, not 0.01
    MsgBox curCommission          ' every figure is 10x too small
End Sub
