VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RibbonModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ModuleName    As String = "RibbonModel."
Private Const MsgBoxTitle   As String = "Rubberduck Ribbon Demo"

Private Type TRibbonModel
    BrandingGroup   As RibbonGroup
    CustomGroup     As RibbonGroup
    StandardGroup   As RibbonGroup
End Type
Private this As TRibbonModel

Private WithEvents mCustomButton1       As RibbonButton
Attribute mCustomButton1.VB_VarHelpID = -1
Private WithEvents mCustomButton2       As RibbonButton
Attribute mCustomButton2.VB_VarHelpID = -1
Private WithEvents mCustomButton3       As RibbonButton
Attribute mCustomButton3.VB_VarHelpID = -1
Private WithEvents mCustomSizeToggle    As RibbonToggleButton
Attribute mCustomSizeToggle.VB_VarHelpID = -1
Private WithEvents mCustomOptions       As RibbonDropDown
Attribute mCustomOptions.VB_VarHelpID = -1

Private WithEvents mBrandingButton      As RibbonButton
Attribute mBrandingButton.VB_VarHelpID = -1
Private WithEvents mStandardButton1     As RibbonButton
Attribute mStandardButton1.VB_VarHelpID = -1
Private WithEvents mStandardButton2     As RibbonButton
Attribute mStandardButton2.VB_VarHelpID = -1
Private WithEvents mStandardOptions     As RibbonDropDown
Attribute mStandardOptions.VB_VarHelpID = -1
Private WithEvents mShowAdvancedToggle  As RibbonToggleButton
Attribute mShowAdvancedToggle.VB_VarHelpID = -1

Public Function Initialize(ByVal Factory As IRibbonFactory) As RibbonModel
    On Error GoTo EH
    Set this.BrandingGroup = Factory.NewRibbonGroup("BrandingGroup")
    Set mBrandingButton = Factory.NewRibbonButtonMso("BrandingButton", , , , "MacroSecurity", True)
    
    Set this.CustomGroup = Factory.NewRibbonGroup("CustomButtonsGroup", False)
    Set mCustomButton1 = Factory.NewRibbonButtonMso("AppLaunchButton1", True, True, rdLarge, "RefreshAll", True, True)
    Set mCustomButton2 = Factory.NewRibbonButtonMso("AppLaunchButton2", True, True, rdLarge, "Refresh", True, True)
    Set mCustomButton3 = Factory.NewRibbonButtonMso("AppLaunchButton3", True, True, rdLarge, "MacroPlay", True, True)
    Set mCustomSizeToggle = Factory.NewRibbonToggleMso("SizeToggle", True, True, rdLarge, DefaultImage, False, True)
    Set mCustomOptions = Factory.NewRibbonDropDown("ButtonOptions2", True, False)
    mCustomOptions.AddItem Factory.NewSelectableItem("LabelOnly")
    mCustomOptions.AddItem Factory.NewSelectableItem("ImageOnly")
    mCustomOptions.AddItem Factory.NewSelectableItem("LabelAndImage")
    
    Set this.StandardGroup = Factory.NewRibbonGroup("StandardButtonsGroup")
    Set mStandardButton1 = Factory.NewRibbonButtonMso("StandardButton1", True, True, rdRegular, "RefreshAll")
    Set mStandardButton2 = Factory.NewRibbonButtonMso("StandardButton2", True, True, rdRegular, "Refresh")
    Set mStandardOptions = Factory.NewRibbonDropDown("ButtonOptions", True, True)
    Set mShowAdvancedToggle = Factory.NewRibbonToggle("ShowAdvancedToggle", True, True)
    
    mShowAdvancedToggle.OnToggled False
    mStandardOptions.OnActionDropDown "", 2

    mCustomOptions.SelectedItemId = "LabelAndImage"
    mCustomOptions.IsEnabled = mCustomSizeToggle.IsPressed
    
    Set Initialize = Me
XT: Exit Function
EH: ReraiseError Err, ModuleName & "Initialize"
    Resume XT
    Resume Next
End Function

Private Sub mBrandingButton_Clicked()
    MsgBox "Quack!", vbOKOnly Or vbInformation, "Rubberduck Ribbon Refresh"
End Sub

Private Sub mCustomButton1_Clicked()
    DefaultButtonAction mCustomButton1
End Sub
Private Sub mCustomButton2_Clicked()
    DefaultButtonAction mCustomButton2
End Sub
Private Sub mCustomButton3_Clicked()
    DefaultButtonAction mCustomButton3
End Sub
Private Sub mCustomOptions_SelectionMade(ByVal ItemId As String, ByVal ItemIndex As Long)
    SetButtonView ItemIndex, mCustomButton1, mCustomButton2, mCustomButton3
End Sub

Private Sub mCustomSizeToggle_Toggled(ByVal IsPressed As Boolean)
    ToggleCustomSize Not IsPressed, mCustomButton1, mCustomButton2, mCustomButton3
End Sub
Private Sub ToggleCustomSize(ByVal IsLarge As Boolean, ParamArray Buttons())
    Dim i As Long
    For i = LBound(Buttons) To UBound(Buttons)
        Buttons(i).Size = IIf(IsLarge, rdLarge, rdRegular)
    Next i
    mCustomOptions.IsEnabled = Not IsLarge
End Sub

Private Sub mStandardButton1_Clicked()
    DefaultButtonAction mStandardButton1
End Sub
Private Sub mStandardButton2_Clicked()
    DefaultButtonAction mStandardButton2
End Sub
Private Sub mStandardOptions_SelectionMade(ByVal ItemId As String, ByVal ItemIndex As Long)
    SetButtonView ItemIndex, mStandardButton1, mStandardButton2
End Sub
Private Sub mShowAdvancedToggle_Toggled(ByVal IsPressed As Boolean)
    this.CustomGroup.IsVisible = IsPressed
End Sub

Private Sub DefaultButtonAction(ByVal Button As RibbonButton)
    MsgBox Button.Id & " Pressed", vbOKOnly Or vbInformation, MsgBoxTitle
End Sub

Private Sub SetButtonView(ByVal SelectedIndex As Integer, ParamArray Buttons())
    Dim i As Long
    Dim ShowLabel As Boolean: ShowLabel = ((SelectedIndex + 1) And 1) <> 0
    Dim ShowImage As Boolean: ShowImage = ((SelectedIndex + 1) And 2) <> 0
    For i = LBound(Buttons) To UBound(Buttons)
        Buttons(i).ShowLabel = ShowLabel
        Buttons(i).ShowImage = ShowImage
    Next i
End Sub
