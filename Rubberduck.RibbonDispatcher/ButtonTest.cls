VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ButtonTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This Visual Basic 6.0 client creates an instance of the Button class and
' implements the event sink interface. The WithEvents directive
' registers the sink interface pointer with the source.
Public WithEvents myButton As ComEventTest.Button
Attribute myButton.VB_VarHelpID = -1

Private Sub Class_Initialize()
    With New ComEventTest.Master
        Set myButton = .NewButton
    End With
End Sub
' Events and methods are matched by name and signature.
Private Sub myButton_Click(ByVal x As Long, ByVal y As Long)
    MsgBox "Click event"
End Sub
Private Sub myButton_Pulse()
    MsgBox "Pulse event"
End Sub
Private Sub myButton_Resize()
    MsgBox "Resize event"
End Sub
