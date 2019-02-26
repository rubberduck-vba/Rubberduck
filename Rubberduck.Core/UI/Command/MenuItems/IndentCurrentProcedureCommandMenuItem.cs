using Rubberduck.UI.Command.ComCommands;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class IndentCurrentProcedureCommandMenuItem : CommandMenuItemBase
    {
        public IndentCurrentProcedureCommandMenuItem(IndentCurrentProcedureCommand command) : base(command)
        {
        }

        public override string Key => "IndentCurrentProcedure";
        public override int DisplayOrder => (int)SmartIndenterMenuItemDisplayOrder.CurrentProcedure;
    }
}
