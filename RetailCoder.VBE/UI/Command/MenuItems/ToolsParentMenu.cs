using System.Collections.Generic;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class ToolsParentMenu : ParentMenuItemBase
    {
        public ToolsParentMenu(IEnumerable<IMenuItem> items)
            : base("ToolsMenu", items)
        {
        }

        public override int DisplayOrder
        {
            get
            {
                return (int)RubberduckMenuItemDisplayOrder.Tools;
            }
        }
        
         
    }

    public enum ToolsMenuItemDisplayOrder
    {
        SourceControl,
        ToDoExplorer,
        RegexAssistant,
    }
}
