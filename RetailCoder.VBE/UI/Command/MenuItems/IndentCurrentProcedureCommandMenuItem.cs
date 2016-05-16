using System.Windows.Input;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class IndentCurrentProcedureCommandMenuItem : CommandMenuItemBase
    {
        public IndentCurrentProcedureCommandMenuItem(ICommand command) : base(command)
        {
        }

        public override string Key { get { return "IndentCurrentProcedure"; } }
        public override int DisplayOrder  { get { return (int)SmartIndenterMenuItemDisplayOrder.CurrentProcedure; } }
    }
}
