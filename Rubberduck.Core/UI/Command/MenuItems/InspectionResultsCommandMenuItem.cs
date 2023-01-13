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


    public class WindowsInspectionResultsCommandMenuItem : InspectionResultsCommandMenuItem
    {
        public WindowsInspectionResultsCommandMenuItem(InspectionResultsCommand command)
            : base(command)
        { }

        public override int DisplayOrder => (int)WindowMenuItemDisplayOrder.CodeInspections;
    }
}
