using System.Windows.Input;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class IndentCurrentModuleCommandMenuItem : CommandMenuItemBase
    {
        public IndentCurrentModuleCommandMenuItem(ICommand command)
            : base(command)
        {
        }

        public override string Key { get { return "IndentCurrentModule"; } }
        public override int DisplayOrder { get { return (int)SmartIndenterMenuItemDisplayOrder.CurrentModule; } }
    }
}
