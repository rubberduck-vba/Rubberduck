using Rubberduck.VBEditor.SafeComWrappers.VB.Enums;
using Rubberduck.VBEditor.SafeComWrappers.Office.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.Office.v12;

namespace Rubberduck.UI.Command.MenuItems.ParentMenus
{
    internal static class CommandBarPopupFactory
    {
        public static ICommandBarPopup Create<TParent>(TParent parent, int? beforeIndex = null)
            where TParent : ICommandBarControls
        {
            return CommandBarPopup.FromCommandBarControl(beforeIndex.HasValue
                ? parent.Add(ControlType.Popup, beforeIndex.Value)
                : parent.Add(ControlType.Popup));
        }
    }
}