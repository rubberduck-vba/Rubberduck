using Rubberduck.VBEditor.SafeComWrappers.MSForms;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract;

namespace Rubberduck.UI.Command.MenuItems.ParentMenus
{
    internal static class CommandBarButtonFactory
    {
        public static ICommandBarButton Create<TParent>(TParent parent, int? beforeIndex = null)
            where TParent : ICommandBarControls
        {
            return CommandBarButton.FromCommandBarControl(beforeIndex.HasValue
                ? parent.Add(ControlType.Button, beforeIndex.Value)
                : parent.Add(ControlType.Button));
        }
    }
}