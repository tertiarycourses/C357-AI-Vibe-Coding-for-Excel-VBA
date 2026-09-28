VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    MsgBox "Welcome to the NorthWind Commission Tracker"
End Sub


Private Sub Workbook_BeforeClose(Cancel As Boolean)

    ' Block the close until the month's commission rate has been entered
    If Worksheets("Commission").Range("F5").Value = "" Then
        MsgBox "Please enter a commission rate in cell F5 " & _
               "before closing the workbook."
        Cancel = True
    End If

End Sub
