namespace Rubberduck.Parsing.Symbols
{
    /*
     *  todo: implement, then remove #region MsFormsLib from VbaStandardLib class.
     */

    //public class MsFormsLib
    //{
    //    private static IEnumerable<Declaration> _msFormsLibDeclarations;

    //    public static IEnumerable<Declaration> Declarations
    //    {
    //        get
    //        {
    //            if (_msFormsLibDeclarations == null)
    //            {
    //                var nestedTypes = typeof(VbaStandardLib).GetNestedTypes(BindingFlags.NonPublic);
    //                var fields = nestedTypes.SelectMany(t => t.GetFields());
    //                var values = fields.Select(f => f.GetValue(null));
    //                _msFormsLibDeclarations = values.Cast<Declaration>();
    //            }

    //            return _msFormsLibDeclarations;
    //        }
    //    }

    //    private static readonly QualifiedModuleName MsFormsModuleName = new QualifiedModuleName("MSForms", "MSForms");

    //    private class UserFormClass
    //    {
    //        public static Declaration UserForm = new Declaration(new QualifiedMemberName(MsFormsModuleName, "UserForm"), "MSForms", "UserForm", true, false, Accessibility.Global, DeclarationType.Class);

    //        // events
    //        public static Declaration AddControl = new Declaration(new QualifiedMemberName(MsFormsModuleName, "AddControl"), "MSForms.Control", null, true, false, Accessibility.Public, DeclarationType.Event);
    //        public static Declaration BeforeDragOver = new Declaration(new QualifiedMemberName(MsFormsModuleName, "BeforeDragOver"), "MSForms.Control", null, true, false, Accessibility.Public, DeclarationType.Event);
    //        public static Declaration BeforeDropOrPaste = new Declaration(new QualifiedMemberName(MsFormsModuleName, "BeforeDropOrPaste"), "MSForms.Control", null, true, false, Accessibility.Public, DeclarationType.Event);
    //        public static Declaration Click = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Click"), "MSForms.Control", null, true, false, Accessibility.Public, DeclarationType.Event);
    //        public static Declaration DblClick = new Declaration(new QualifiedMemberName(MsFormsModuleName, "DblClick"), "MSForms.Control", null, true, false, Accessibility.Public, DeclarationType.Event);
    //        public static Declaration Error = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Error"), "MSForms.Control", null, true, false, Accessibility.Public, DeclarationType.Event);
    //        public static Declaration KeyDown = new Declaration(new QualifiedMemberName(MsFormsModuleName, "KeyDown"), "MSForms.Control", null, true, false, Accessibility.Public, DeclarationType.Event);
    //        public static Declaration KeyPress = new Declaration(new QualifiedMemberName(MsFormsModuleName, "KeyPress"), "MSForms.Control", null, true, false, Accessibility.Public, DeclarationType.Event);
    //        public static Declaration KeyUp = new Declaration(new QualifiedMemberName(MsFormsModuleName, "KeyUp"), "MSForms.Control", null, true, false, Accessibility.Public, DeclarationType.Event);
    //        public static Declaration Layout = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Layout"), "MSForms.Control", null, true, false, Accessibility.Public, DeclarationType.Event);
    //        public static Declaration MouseDown = new Declaration(new QualifiedMemberName(MsFormsModuleName, "MouseDown"), "MSForms.Control", null, true, false, Accessibility.Public, DeclarationType.Event);
    //        public static Declaration MouseMove = new Declaration(new QualifiedMemberName(MsFormsModuleName, "MouseMove"), "MSForms.Control", null, true, false, Accessibility.Public, DeclarationType.Event);
    //        public static Declaration MouseUp = new Declaration(new QualifiedMemberName(MsFormsModuleName, "MouseUp"), "MSForms.Control", null, true, false, Accessibility.Public, DeclarationType.Event);
    //        public static Declaration RemoveControl = new Declaration(new QualifiedMemberName(MsFormsModuleName, "RemoveControl"), "MSForms.Control", null, true, false, Accessibility.Public, DeclarationType.Event);
    //        public static Declaration Scroll = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Scroll"), "MSForms.Control", null, true, false, Accessibility.Public, DeclarationType.Event);
    //        public static Declaration Zoom = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Zoom"), "MSForms.Control", null, true, false, Accessibility.Public, DeclarationType.Event);

    //        // ghost events (nowhere in the object browser)
    //        public static Declaration Activate = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Activate"), "MSForms.Control", null, true, false, Accessibility.Public, DeclarationType.Event);
    //        public static Declaration Deactivate = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Deactivate"), "MSForms.Control", null, true, false, Accessibility.Public, DeclarationType.Event);
    //        public static Declaration Initialize = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Initialize"), "MSForms.Control", null, true, false, Accessibility.Public, DeclarationType.Event);
    //        public static Declaration QueryClose = new Declaration(new QualifiedMemberName(MsFormsModuleName, "QueryClose"), "MSForms.Control", null, true, false, Accessibility.Public, DeclarationType.Event);
    //        public static Declaration Resize = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Resize"), "MSForms.Control", null, true, false, Accessibility.Public, DeclarationType.Event);
    //        public static Declaration Terminate = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Terminate"), "MSForms.Control", null, true, false, Accessibility.Public, DeclarationType.Event);
    //    }

    //    private class ControlsClass
    //    {
    //        public static Declaration Controls = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Controls"), "MSForms", "Controls", true, false, Accessibility.Global, DeclarationType.Class);
    //        public static Declaration Count = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Count"), "MSForms.Controls", "Long", true, false, Accessibility.Global, DeclarationType.PropertyGet);
    //        public static Declaration Add = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Add"), "MSForms.Controls", "Control", true, false, Accessibility.Public, DeclarationType.Function);
    //        public static Declaration AddByClass = new Declaration(new QualifiedMemberName(MsFormsModuleName, "[_AddByClass]"), "MSForms.Controls", "Control", true, false, Accessibility.Public, DeclarationType.Function);
    //        public static Declaration AlignToGrid = new Declaration(new QualifiedMemberName(MsFormsModuleName, "AlignToGrid"), "MSForms.Controls", null, true, false, Accessibility.Public, DeclarationType.Procedure);
    //        public static Declaration BringForward = new Declaration(new QualifiedMemberName(MsFormsModuleName, "BringForward"), "MSForms.Controls", null, true, false, Accessibility.Public, DeclarationType.Procedure);
    //        public static Declaration BringToFront = new Declaration(new QualifiedMemberName(MsFormsModuleName, "BringToFront"), "MSForms.Controls", null, true, false, Accessibility.Public, DeclarationType.Procedure);
    //        public static Declaration Clear = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Clear"), "MSForms.Controls", null, true, false, Accessibility.Public, DeclarationType.Procedure);
    //        public static Declaration Copy = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Copy"), "MSForms.Controls", null, true, false, Accessibility.Public, DeclarationType.Procedure);
    //        public static Declaration Cut = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Cut"), "MSForms.Controls", null, true, false, Accessibility.Public, DeclarationType.Procedure);
    //        public static Declaration Enum = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Enum"), "MSForms.Controls", "Unknown", true, false, Accessibility.Public, DeclarationType.Function);
    //        public static Declaration GetItemById = new Declaration(new QualifiedMemberName(MsFormsModuleName, "[_GetItemByID]"), "MSForms.Controls", "Control", true, false, Accessibility.Public, DeclarationType.Function);
    //        public static Declaration GetItemByIndex = new Declaration(new QualifiedMemberName(MsFormsModuleName, "[_GetItemByIndex]"), "MSForms.Controls", "Control", true, false, Accessibility.Public, DeclarationType.Function);
    //        public static Declaration GetItemByName = new Declaration(new QualifiedMemberName(MsFormsModuleName, "[_GetItemByName]"), "MSForms.Controls", "Control", true, false, Accessibility.Public, DeclarationType.Function);
    //        public static Declaration NewEnum = new Declaration(new QualifiedMemberName(MsFormsModuleName, "[_NewEnum]"), "VBA.Collection", "Unknown", false, false, Accessibility.Public, DeclarationType.PropertyGet);
    //        public static Declaration Item = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Item"), "MSForms.Controls", "Object", true, false, Accessibility.Public, DeclarationType.Function);
    //        public static Declaration Move = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Move"), "MSForms.Controls", null, true, false, Accessibility.Public, DeclarationType.Procedure);
    //        public static Declaration _Move = new Declaration(new QualifiedMemberName(MsFormsModuleName, "[_Move]"), "MSForms.Controls", null, true, false, Accessibility.Public, DeclarationType.Procedure);
    //        public static Declaration Remove = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Remove"), "MSForms.Controls", null, true, false, Accessibility.Public, DeclarationType.Procedure);
    //        public static Declaration SelectAll = new Declaration(new QualifiedMemberName(MsFormsModuleName, "SelectAll"), "MSForms.Controls", null, true, false, Accessibility.Public, DeclarationType.Procedure);
    //        public static Declaration SendBackward = new Declaration(new QualifiedMemberName(MsFormsModuleName, "SendBackward"), "MSForms.Controls", null, true, false, Accessibility.Public, DeclarationType.Procedure);
    //        public static Declaration SendToBack = new Declaration(new QualifiedMemberName(MsFormsModuleName, "SendToBack"), "MSForms.Controls", null, true, false, Accessibility.Public, DeclarationType.Procedure);
    //    }

    //    private class ControlClass
    //    {
    //        public static Declaration Control = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Control"), "MSForms", "Control", true, false, Accessibility.Global, DeclarationType.Class);

    //        // properties
    //        public static Declaration GetCancel = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Cancel"), "MSForms.Control", "Boolean", true, false, Accessibility.Global, DeclarationType.PropertyGet);
    //        public static Declaration LetCancel = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Cancel"), "MSForms.Control", "Boolean", true, false, Accessibility.Global, DeclarationType.PropertyLet);
    //        public static Declaration GetControlSource = new Declaration(new QualifiedMemberName(MsFormsModuleName, "ControlSource"), "MSForms.Control", "String", true, false, Accessibility.Global, DeclarationType.PropertyGet);
    //        public static Declaration LetControlSource = new Declaration(new QualifiedMemberName(MsFormsModuleName, "ControlSource"), "MSForms.Control", "String", true, false, Accessibility.Global, DeclarationType.PropertyLet);
    //        public static Declaration GetControlTipText = new Declaration(new QualifiedMemberName(MsFormsModuleName, "ControlTipText"), "MSForms.Control", "String", true, false, Accessibility.Global, DeclarationType.PropertyGet);
    //        public static Declaration LetControlTipText = new Declaration(new QualifiedMemberName(MsFormsModuleName, "ControlTipText"), "MSForms.Control", "String", true, false, Accessibility.Global, DeclarationType.PropertyLet);
    //        public static Declaration GetDefault = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Default"), "MSForms.Control", "Boolean", true, false, Accessibility.Global, DeclarationType.PropertyGet);
    //        public static Declaration LetDefault = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Default"), "MSForms.Control", "Boolean", true, false, Accessibility.Global, DeclarationType.PropertyLet);
    //        public static Declaration GetHeight = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Height"), "MSForms.Control", "Single", true, false, Accessibility.Global, DeclarationType.PropertyGet);
    //        public static Declaration LetHeight = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Height"), "MSForms.Control", "Single", true, false, Accessibility.Global, DeclarationType.PropertyLet);
    //        public static Declaration GetHelpContextId = new Declaration(new QualifiedMemberName(MsFormsModuleName, "HelpContextID"), "MSForms.Control", "Long", true, false, Accessibility.Global, DeclarationType.PropertyGet);
    //        public static Declaration LetHelpContextId = new Declaration(new QualifiedMemberName(MsFormsModuleName, "HelpContextID"), "MSForms.Control", "Long", true, false, Accessibility.Global, DeclarationType.PropertyLet);
    //        public static Declaration LayoutEffect = new Declaration(new QualifiedMemberName(MsFormsModuleName, "LayoutEffect"), "MSForms.Control", "fmLayoutEffect", true, false, Accessibility.Global, DeclarationType.PropertyGet);
    //        public static Declaration GetLeft = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Left"), "MSForms.Control", "Single", true, false, Accessibility.Global, DeclarationType.PropertyGet);
    //        public static Declaration LetLeft = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Left"), "MSForms.Control", "Single", true, false, Accessibility.Global, DeclarationType.PropertyLet);
    //        public static Declaration GetName = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Name"), "MSForms.Control", "String", true, false, Accessibility.Global, DeclarationType.PropertyGet);
    //        public static Declaration LetName = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Name"), "MSForms.Control", "String", true, false, Accessibility.Global, DeclarationType.PropertyLet);
    //        public static Declaration Object = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Object"), "MSForms.Control", "Object", true, false, Accessibility.Global, DeclarationType.PropertyGet);
    //        public static Declaration OldHeight = new Declaration(new QualifiedMemberName(MsFormsModuleName, "OldHeight"), "MSForms.Control", "Single", true, false, Accessibility.Global, DeclarationType.PropertyGet);
    //        public static Declaration OldLeft = new Declaration(new QualifiedMemberName(MsFormsModuleName, "OldLeft"), "MSForms.Control", "Single", true, false, Accessibility.Global, DeclarationType.PropertyGet);
    //        public static Declaration OldTop = new Declaration(new QualifiedMemberName(MsFormsModuleName, "OldTop"), "MSForms.Control", "Single", true, false, Accessibility.Global, DeclarationType.PropertyGet);
    //        public static Declaration OldWidth = new Declaration(new QualifiedMemberName(MsFormsModuleName, "OldWidth"), "MSForms.Control", "Single", true, false, Accessibility.Global, DeclarationType.PropertyGet);
    //        public static Declaration Parent = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Parent"), "MSForms.Control", "Object", true, false, Accessibility.Global, DeclarationType.PropertyGet);
    //        public static Declaration GetRowSource = new Declaration(new QualifiedMemberName(MsFormsModuleName, "RowSource"), "MSForms.Control", "String", true, false, Accessibility.Global, DeclarationType.PropertyGet);
    //        public static Declaration LetRowSource = new Declaration(new QualifiedMemberName(MsFormsModuleName, "RowSource"), "MSForms.Control", "String", true, false, Accessibility.Global, DeclarationType.PropertyLet);
    //        public static Declaration GetRowSourceType = new Declaration(new QualifiedMemberName(MsFormsModuleName, "RowSourceType"), "MSForms.Control", "Integer", true, false, Accessibility.Global, DeclarationType.PropertyGet);
    //        public static Declaration LetRowSourceType = new Declaration(new QualifiedMemberName(MsFormsModuleName, "RowSourceType"), "MSForms.Control", "Integer", true, false, Accessibility.Global, DeclarationType.PropertyLet);
    //        public static Declaration GetTabIndex = new Declaration(new QualifiedMemberName(MsFormsModuleName, "TabIndex"), "MSForms.Control", "Integer", true, false, Accessibility.Global, DeclarationType.PropertyGet);
    //        public static Declaration LetTabIndex = new Declaration(new QualifiedMemberName(MsFormsModuleName, "TabIndex"), "MSForms.Control", "Integer", true, false, Accessibility.Global, DeclarationType.PropertyLet);
    //        public static Declaration GetTabStop = new Declaration(new QualifiedMemberName(MsFormsModuleName, "TabStop"), "MSForms.Control", "Boolean", true, false, Accessibility.Global, DeclarationType.PropertyGet);
    //        public static Declaration LetTabStop = new Declaration(new QualifiedMemberName(MsFormsModuleName, "TabStop"), "MSForms.Control", "Boolean", true, false, Accessibility.Global, DeclarationType.PropertyLet);
    //        public static Declaration GetTag = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Tag"), "MSForms.Control", "String", true, false, Accessibility.Global, DeclarationType.PropertyGet);
    //        public static Declaration LetTag = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Tag"), "MSForms.Control", "String", true, false, Accessibility.Global, DeclarationType.PropertyLet);
    //        public static Declaration GetTop = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Top"), "MSForms.Control", "Single", true, false, Accessibility.Global, DeclarationType.PropertyGet);
    //        public static Declaration LetTop = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Top"), "MSForms.Control", "Single", true, false, Accessibility.Global, DeclarationType.PropertyLet);
    //        public static Declaration GetVisible = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Visible"), "MSForms.Control", "Boolean", true, false, Accessibility.Global, DeclarationType.PropertyGet);
    //        public static Declaration LetVisible = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Visible"), "MSForms.Control", "Boolean", true, false, Accessibility.Global, DeclarationType.PropertyLet);
    //        public static Declaration GetWidth = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Width"), "MSForms.Control", "Single", true, false, Accessibility.Global, DeclarationType.PropertyGet);
    //        public static Declaration LetWidth = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Width"), "MSForms.Control", "Single", true, false, Accessibility.Global, DeclarationType.PropertyLet);

    //        // procedures
    //        public static Declaration Move = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Move"), "MSForms.Control", null, true, false, Accessibility.Public, DeclarationType.Procedure);
    //        public static Declaration SetFocus = new Declaration(new QualifiedMemberName(MsFormsModuleName, "SetFocus"), "MSForms.Control", null, true, false, Accessibility.Public, DeclarationType.Procedure);
    //        public static Declaration ZOrder = new Declaration(new QualifiedMemberName(MsFormsModuleName, "ZOrder"), "MSForms.Control", null, true, false, Accessibility.Public, DeclarationType.Procedure);

    //        // events
    //        public static Declaration AfterUpdate = new Declaration(new QualifiedMemberName(MsFormsModuleName, "AfterUpdate"), "MSForms.Control", null, true, false, Accessibility.Public, DeclarationType.Event);
    //        public static Declaration BeforeUpdate = new Declaration(new QualifiedMemberName(MsFormsModuleName, "BeforeUpdate"), "MSForms.Control", null, true, false, Accessibility.Public, DeclarationType.Event);
    //        public static Declaration Enter = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Enter"), "MSForms.Control", null, true, false, Accessibility.Public, DeclarationType.Event);
    //        public static Declaration Exit = new Declaration(new QualifiedMemberName(MsFormsModuleName, "Exit"), "MSForms.Control", null, true, false, Accessibility.Public, DeclarationType.Event);
    //    }
    //}
}
