using System.Collections.Generic;

namespace Rubberduck.UI.Command.MenuItems.ParentMenus
{
    public class ToolsParentMenu : ParentMenuItemBase
    {
        public ToolsParentMenu(IEnumerable<IMenuItem> items)
            : base("ToolsMenu", items)
        {
        }

        public override int DisplayOrder => (int)RubberduckMenuItemDisplayOrder.Tools;
    }

    public enum ToolsMenuItemDisplayOrder
    {
        CodeMetrics,
        ToDoExplorer,
        RegexAssistant,
        ExportAll,
        AddRemoveReferences
    }
}
