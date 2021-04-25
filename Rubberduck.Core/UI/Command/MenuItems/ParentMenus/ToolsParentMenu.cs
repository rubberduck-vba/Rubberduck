using System.Collections.Generic;
using Rubberduck.Parsing.UIContext;

namespace Rubberduck.UI.Command.MenuItems.ParentMenus
{
    public class ToolsParentMenu : ParentMenuItemBase
    {
        public ToolsParentMenu(IEnumerable<IMenuItem> items, IUiDispatcher dispatcher)
            : base(dispatcher, "ToolsMenu", items)
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
