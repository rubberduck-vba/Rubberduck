namespace Rubberduck.VBEditor.SafeComWrappers
{
    // Abstraction of the MsoControlType enum in the interop assemblies for Office.v8 and Office.v12
    public enum ControlType
    {
        Custom = 0,
        Button = 1,
        Edit = 2,
        Dropdown = 3,
        ComboBox = 4,
        ButtonDropdown = 5,
        SplitDropdown = 6,
        OcxDropdown = 7,
        GenericDropdown = 8,
        GraphicDropdown = 9,
        Popup = 10,
        GraphicPopup = 11,
        ButtonPopup = 12,
        SplitButtonPopup = 13,
        SplitButtonMruPopup = 14,
        Label = 15,
        ExpandingGrid = 16,
        SplitExpandingGrid = 17,
        Grid = 18,
        Gauge = 19,
        GraphicCombo = 20,
        Pane = 21,                  // Not in Office.v8
        ActiveX = 22,               // Not in Office.v8
        Spinner = 23,               // Not in Office.v8
        LabelEx = 24,               // Not in Office.v8
        WorkPane = 25,              // Not in Office.v8
        AutoCompleteCombo = 26,     // Not in Office.v8
    }
}
