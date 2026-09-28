Attribute VB_Name = "Module1"

Sub Welcome()
    ' Greets the analyst when the commission tool is opened
    MsgBox "Welcome to the NorthWind Commission Tracker"
End Sub


Sub ConfirmRecalculate()

    Dim intAnswer As Integer

    intAnswer = MsgBox("Recalculate all commissions?", _
                       vbYesNoCancel, "Confirm")

    If intAnswer = vbYes Then
        MsgBox "Recalculating..."
    ElseIf intAnswer = vbNo Then
        MsgBox "Cancelled by user."
    Else
        MsgBox "Operation aborted."
    End If

End Sub


Sub ShowCommission()

    Dim curRate As Double
    Dim curSale As Currency
    Dim curCommission As Currency

    curRate = InputBox("What is the commission rate? (e.g. 0.1 for 10%)")
    curSale = ActiveCell.Value
    curCommission = curSale * curRate

    MsgBox "Commission: " & Format(curCommission, "$#,##0.00")

End Sub
