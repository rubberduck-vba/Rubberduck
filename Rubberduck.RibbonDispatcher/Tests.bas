Attribute VB_Name = "Tests"
Option Explicit

Public Sub Test()
    On Error GoTo EH
    With New ButtonTest
        .myButton.CauseClickEvent 400, 600
        .myButton.CausePulse
        .myButton.CauseResizeEvent
    End With
XT: Exit Sub
EH: MsgBox "Error #" & Err.Number & " - " & Err.Description, _
    vbOKOnly Or vbExclamation, "Button Test"
    Resume Next
End Sub
