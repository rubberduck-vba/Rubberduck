using Rubberduck.UI.Command.ComCommands;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class InspectionResultsCommandMenuItem : CommandMenuItemBase
    {
        public InspectionResultsCommandMenuItem(InspectionResultsCommand command) 
            : base(command)
        {
        }

        public override string Key => "RubberduckMenu_CodeInspections";
        public override int DisplayOrder => (int)RubberduckMenuItemDisplayOrder.CodeInspections;
    }
}
