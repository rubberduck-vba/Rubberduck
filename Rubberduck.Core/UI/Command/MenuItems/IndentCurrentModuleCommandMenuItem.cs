using Rubberduck.UI.Command.ComCommands;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class IndentCurrentModuleCommandMenuItem : CommandMenuItemBase
    {
        public IndentCurrentModuleCommandMenuItem(IndentCurrentModuleCommand command)
            : base(command)
        {
        }

        public override string Key => "IndentCurrentModule";
        public override int DisplayOrder => (int)SmartIndenterMenuItemDisplayOrder.CurrentModule;
    }
}
