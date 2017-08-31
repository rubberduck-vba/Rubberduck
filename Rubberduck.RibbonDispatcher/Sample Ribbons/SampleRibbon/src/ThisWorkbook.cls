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

Private Const ModuleName As String = "ThisWorkbook."

Private Type TRIbbonWorkbook
    ViewModel      As Rubberduck_RibbonDispatcher.RibbonViewModel
    RibbonModel    As RibbonModel
End Type
Private this As TRIbbonWorkbook

Implements Rubberduck_RibbonDispatcher.IRibbonWorkbook

Friend Property Get IRibbonWorkbook_ViewModel() As Rubberduck_RibbonDispatcher.RibbonViewModel
    Set IRibbonWorkbook_ViewModel = this.ViewModel
End Property

Friend Sub InitializeRibbonViewModel(ByVal RibbonUI As Office.IRibbonUI)
    On Error GoTo EH
    MsgBox "Pause to Ctrl-Break", vbOKOnly Or vbInformation, "InitializeRibbonViewmodel"
    With New Rubberduck_RibbonDispatcher.Main
        Set this.ViewModel = .NewRibbonViewModel(RibbonUI, New MyResourceManager)
    End With
    With New RibbonModel
        Set this.RibbonModel = .Initialize(this.ViewModel.RibbonFactory)
    End With
    RibbonUI.ActivateTab "TabRubberDuck"

XT: Exit Sub
EH: DisplayError Err, ModuleName & "OnRibbonLoad"
    Resume XT
End Sub

Private Sub Workbook_Open()
    MsgBox "Pause to Ctrl-Break", vbOKOnly Or vbInformation, "Workbook_Open"
End Sub
