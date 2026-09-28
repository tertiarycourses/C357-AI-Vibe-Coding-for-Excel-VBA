VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet_Commission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' Paste this code into the 'Commission' worksheet's code window
' (right-click the sheet tab -> View Code).

Private Sub Worksheet_Change(ByVal Target As Range)

    Dim intResponse As Integer

    intResponse = MsgBox("This is an automatically generated " & _
                         "worksheet. " & _
                         "Are you sure you want to change it?", vbOKCancel)

    If intResponse = vbOK Then
        MsgBox "Worksheet value changed."
    Else
        ' stop the undo from re-triggering this same event
        Application.EnableEvents = False
        Application.Undo
        Application.EnableEvents = True
        MsgBox "Change rejected."
    End If

End Sub
