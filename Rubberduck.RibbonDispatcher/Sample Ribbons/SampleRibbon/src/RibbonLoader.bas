Attribute VB_Name = "RibbonLoader"
Option Explicit

Public Sub OnRibbonLoad(ByVal RibbonUI As Office.IRibbonUI)
    On Error GoTo EH
    ThisWorkbook.InitializeRibbonViewModel RibbonUI
XT: Exit Sub
EH: DisplayError Err, "RibbonLoader.OnRibbonLoad"
    Resume XT
End Sub
