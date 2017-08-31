VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyResourceManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ModuleName As String = "MyResourceManager."

Private Type TRibbonStrings
    Strings As Rubberduck_RibbonDispatcher.ControlStrings
End Type
Private this As TRibbonStrings

Implements Rubberduck_RibbonDispatcher.IResourceManager

Private Function IResourceManager_GetCurrentUIString(ByVal Name As String) As String
    On Error GoTo EH
    IResourceManager_GetCurrentUIString = this.Strings.Item(Name)

XT: Exit Function
EH: ReraiseError Err, ModuleName & "IResourceManager_GetCurrentUIString"
    Resume XT
End Function

Private Function IResourceManager_LoadImage(ByVal ImageId As String)
    On Error GoTo EH
    IResourceManager_LoadImage = Nothing
XT: Exit Function
EH: ReraiseError Err, ModuleName & "IResourceManager_LoadImage"
    Resume XT
End Function

Private Sub Class_Initialize()
    On Error GoTo EH
    Set this.Strings = New Rubberduck_RibbonDispatcher.ControlStrings
    With this.Strings
        .AddControl "BrandingGroup", "RubberDuck", _
            "Tools for VBA Development"
        
        .AddControl "BrandingButton", "RubberDuck", _
            "Tools for VBA Development", _
            "A library of tools for VBA development, including utilities for source-code control and unit-testing."
        
        .AddControl "StandardButtonsGroup", "Standard Buttons", _
            "Buttons Always Visible", _
            "These buttons are always visible on this ribbon."
        
        .AddControl "StandardButton1", "1st Std Button", _
            "First Std Button", _
            "Demonstration of a 'standard button'; one that is always visible on the ribbon."
        
        .AddControl "StandardButton2", "2nd Std Button", _
            "Second Std Button", _
            "Another demonstration of a 'standard button'; one that is always visible on the ribbon."
        
        .AddControl "ButtonOptions", "Options", _
            "Controls Button Display", _
            "Sets display options for standard buttons."
        
        .AddControl "ShowAdvancedToggle", "Show Advanced", _
            "Toggles Display of Advanced Buttons", _
            "Alternates between setting the advanced buttons visible and hidden.", ""
        
        .AddControl "CustomButtonsGroup", "Custom Buttons", _
            "Buttons Optionally Visible", _
            "These buttons are visible on this ribbon only when requested."
        
        .AddControl "AppLaunchButton1", "1st Custom Button", _
            "First Custom Button", "First example of a custom button."
        
        .AddControl "AppLaunchButton2", "2nd Custom Button", _
            "Seccond Custom Button", "Second example of a custom button."
        
        .AddControl "AppLaunchButton3", "3rd Custom Button", _
            "Third Custom Button", "Third example of a custom button."
        
        .AddControl "ButtonOptions2", "Options", _
            "Advanced Button Options", _
            "Display options for Advanced Buttons:" & _
                    vbNewLine & "Label Only;" & _
                    vbNewLine & "Image Only;" & _
                    vbNewLine & "Label & Image"
        
        .AddControl "ImageOnly", "Image Only", _
            "Display Image Only", _
            "Displays just the button image when its size is Regular; else both."
        
        .AddControl "LabelOnly", "Label Only", _
            "Display Label Only", _
            "Displays just the button label when its size is Regular; else both."
        
        .AddControl "LabelAndImage", "Label & Image", _
            "Display Label & Image", _
            "Always displays both the button label and image."
        
        .AddControl "SizeToggle", "Prefer Small", _
            "Toggles Button Size", _
            "Toggles between large and small (preferred size for) buttons for this ribbon group.", _
            "Prefer Large"
    End With
XT: Exit Sub
EH: ReraiseError Err, ModuleName & "Class_Initialize"
    Resume XT
End Sub
