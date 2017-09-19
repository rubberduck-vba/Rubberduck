Attribute VB_Name = "RibbonCallbacksAdapter"
Option Explicit

Public Const ModuleName   As String = "RibbonCallbacksAdapter."
Public Const DefaultImage As String = "MacroSecurity"

Public Enum ControlSize
    rdRegular = RdControlSize_rdRegular
    rdLarge = RdControlSize_rdLarge
End Enum

Private Function ViewModelFor(ByVal Control As IRibbonControl _
) As Rubberduck_RibbonDispatcher.RibbonViewModel
    On Error GoTo EH
    Dim WkBk As IRibbonWorkbook
    Set WkBk = Control.Context.Parent
    Set ViewModelFor = WkBk.ViewModel
XT: Exit Function
EH: If Err.Number <> 438 Then ReraiseError Err, ModuleName & "NewRibbonModel"
    Resume XT
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                             Published Actions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadImage(ByVal ImageId, ByRef Image)
    On Error GoTo EH
    Set Image = Nothing
XT: Exit Sub
EH: If Err.Number <> 91 Then DisplayError Err, ModuleName & "LoadImage", ImageId
    Resume XT
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                         Common to Many Controls
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetEnabled(ByVal Control As IRibbonControl, ByRef Enabled)
    On Error GoTo EH
    Enabled = ViewModelFor(Control).GetEnabled(Control)
XT: Exit Sub
EH: If Err.Number <> 91 Then DisplayError Err, ModuleName & "GetEnabled", Control.Id
    Resume XT
End Sub
Public Sub GetVisible(ByVal Control As IRibbonControl, ByRef Visible)
    On Error GoTo EH
    Visible = ViewModelFor(Control).GetVisible(Control)
XT: Exit Sub
EH: If Err.Number <> 91 Then DisplayError Err, ModuleName & "GetVisible", Control.Id
    Resume XT
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                             Control Text Strings
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetDescription(ByVal Control As IRibbonControl, ByRef Description)
    On Error GoTo EH
    Description = ViewModelFor(Control).GetDescription(Control)
XT: Exit Sub
EH: If Err.Number <> 91 Then DisplayError Err, ModuleName & "GetDescription", Control.Id
    Resume XT
End Sub
Public Sub GetLabel(ByVal Control As IRibbonControl, ByRef Label)
    On Error GoTo EH
    Label = ViewModelFor(Control).GetLabel(Control)
XT: Exit Sub
EH: If Err.Number <> 91 Then DisplayError Err, ModuleName & "GetLabel", Control.Id
    Resume XT
End Sub
Public Sub GetKeytip(ByVal Control As IRibbonControl, ByRef KeyTip)
    On Error GoTo EH
    KeyTip = Mid(ViewModelFor(Control).GetKeytip(Control), 1, 3)
XT: Exit Sub
EH: If Err.Number <> 91 Then DisplayError Err, ModuleName & "GetKeytip", Control.Id
    Resume XT
End Sub
Public Sub GetScreentip(ByVal Control As IRibbonControl, ByRef ScreenTip)
    On Error GoTo EH
    ScreenTip = ViewModelFor(Control).GetScreentip(Control)
XT: Exit Sub
EH: If Err.Number <> 91 Then DisplayError Err, ModuleName & "GetScreentip", Control.Id
    Resume XT
End Sub
Public Sub GetSupertip(ByVal Control As IRibbonControl, ByRef SuperTip)
    On Error GoTo EH
    SuperTip = ViewModelFor(Control).GetSupertip(Control)
XT: Exit Sub
EH: If Err.Number <> 91 Then DisplayError Err, ModuleName & "GetSupertip", Control.Id
    Resume XT
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                             Sizeable Controls
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetSize(ByVal Control As IRibbonControl, ByRef Size)
    On Error GoTo EH
    Size = ViewModelFor(Control).GetSize(Control)
XT: Exit Sub
EH: If Err.Number <> 91 Then DisplayError Err, ModuleName & "GetSize", Control.Id
    Resume XT
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                             Imageable Controls
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetImage(ByVal Control As IRibbonControl, ByRef Image)
    On Error GoTo EH
    Image = ViewModelFor(Control).GetImage(Control)
XT: Exit Sub
EH: If Err.Number <> 91 Then DisplayError Err, ModuleName & "GetImage", Control.Id
    Resume XT
End Sub
Public Sub GetImageMso(ByVal Control As IRibbonControl, ByRef ImageMso)
    On Error GoTo EH
    ImageMso = ViewModelFor(Control).GetImage(Control)
XT: Exit Sub
EH: If Err.Number <> 91 Then DisplayError Err, ModuleName & "GetImageMso", Control.Id
    Resume XT
End Sub
Public Sub GetShowImage(ByVal Control As IRibbonControl, ByRef ShowImage)
    On Error GoTo EH
    ShowImage = ViewModelFor(Control).GetShowImage(Control)
XT: Exit Sub
EH: If Err.Number <> 91 Then DisplayError Err, ModuleName & "GetShowImage", Control.Id
    Resume XT
End Sub
Public Sub GetShowLabel(ByVal Control As IRibbonControl, ByRef ShowLabel)
    On Error GoTo EH
    ShowLabel = ViewModelFor(Control).GetShowLabel(Control)
XT: Exit Sub
EH: If Err.Number <> 91 Then DisplayError Err, ModuleName & "GetShowLabel", Control.Id
    Resume XT
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                     Action Controls - Buttons & Menu Items
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
Public Sub OnAction(ByVal Control As IRibbonControl)
    On Error GoTo EH
    ViewModelFor(Control).OnAction Control
XT: Exit Sub
EH: If Err.Number <> 91 Then DisplayError Err, ModuleName & "OnAction", Control.Id
    Resume XT
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                 Toggle Controls - ToggleButtons & CheckBoxes
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetPressed(ByVal Control As IRibbonControl, ByRef Pressed)
    On Error GoTo EH
    Pressed = ViewModelFor(Control).GetPressed(Control)
XT: Exit Sub
EH: If Err.Number <> 91 Then DisplayError Err, ModuleName & "GetPressed", Control.Id
    Resume XT
End Sub
Public Sub OnActionToggle(ByVal Control As IRibbonControl, ByVal Pressed As Boolean)
    On Error GoTo EH
    ViewModelFor(Control).OnActionToggle Control, Pressed
XT: Exit Sub
EH: If Err.Number <> 91 Then DisplayError Err, ModuleName & "OnActionToggle", Control.Id
    Resume XT
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''             Selectable Contros - DropDowns & ComboBoxes
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub GetItemCount(ByVal Control As IRibbonControl, ByRef Count)
    On Error GoTo EH
    Count = ViewModelFor(Control).GetItemCount(Control)
XT: Exit Sub
EH: If Err.Number <> 91 Then DisplayError Err, ModuleName & "GetItemCount", Control.Id
    Resume XT
End Sub
Sub GetSelectedItemID(ByVal Control As IRibbonControl, ByRef Id)
    On Error GoTo EH
    Id = ViewModelFor(Control).GetSelectedItemID(Control)
XT: Exit Sub
EH: If Err.Number <> 91 Then DisplayError Err, ModuleName & "GetSelectedItemID", Control.Id
    Resume XT
End Sub
Sub GetSelectedItemIndex(ByVal Control As IRibbonControl, ByRef Index)
    On Error GoTo EH
    Index = ViewModelFor(Control).GetSelectedItemIndex(Control)
XT: Exit Sub
EH: If Err.Number <> 91 Then DisplayError Err, ModuleName & "GetSelectedItemIndex", Control.Id
    Resume XT
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''               Selectable Items - DropDown- & ComboBox-Items
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub GetItemID(ByVal Control As IRibbonControl, ByVal Index As Integer, ByRef Id)
    On Error GoTo EH
    Id = ViewModelFor(Control).GetItemID(Control, Index)
XT: Exit Sub
EH: If Err.Number <> 91 Then DisplayError Err, ModuleName & "GetItemID", Control.Id
    Resume XT
End Sub
Sub GetItemImage(ByVal Control As IRibbonControl, ByVal Index As Integer, ByRef Image)
    On Error GoTo EH
    Set Image = ViewModelFor(Control).GetItemImage(Control, Index)
XT: Exit Sub
EH: If Err.Number <> 91 Then Image = "MacroSecurity" ' DisplayError Err, ModuleName & "GetItemImage"
    Resume XT
End Sub
Sub GetItemLabel(ByVal Control As IRibbonControl, ByVal Index As Integer, ByRef Label)
    On Error GoTo EH
    Label = ViewModelFor(Control).GetItemLabel(Control, Index)
XT: Exit Sub
EH: If Err.Number <> 91 Then DisplayError Err, ModuleName & "GetItemLabel", Control.Id
    Resume XT
End Sub
Sub GetItemScreenTip(ByVal Control As IRibbonControl, ByVal Index As Integer, ByRef ScreenTip)
    On Error GoTo EH
    ScreenTip = ViewModelFor(Control).GetItemScreenTip(Control, Index)
XT: Exit Sub
EH: If Err.Number <> 91 Then DisplayError Err, ModuleName & "GetItemScreenTip", Control.Id
    Resume XT
End Sub
Sub GetItemSuperTip(ByVal Control As IRibbonControl, ByVal Index As Integer, ByRef SuperTip)
    On Error GoTo EH
    SuperTip = ViewModelFor(Control).GetItemSuperTip(Control, Index)
XT: Exit Sub
EH: If Err.Number <> 91 Then DisplayError Err, ModuleName & "GetItemSuperTip", Control.Id
    Resume XT
End Sub
Sub OnActionDropDown(ByVal Control As IRibbonControl, ByVal SelectedId As String, ByVal SelectedIndex As Integer)
    On Error GoTo EH
    ViewModelFor(Control).OnActionDropDown Control, SelectedId, SelectedIndex
XT: Exit Sub
EH: If Err.Number <> 91 Then DisplayError Err, ModuleName & "OnActionDropDown", Control.Id
    Resume XT
End Sub
