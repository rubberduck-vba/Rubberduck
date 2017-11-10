using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    class CodeMetricsCommandMenuItem : CommandMenuItemBase
    {
        public CodeMetricsCommandMenuItem(CommandBase command) 
            : base(command)
        {
        }
        public override bool EvaluateCanExecute(RubberduckParserState state) => true;

        public override string Key => "RubberduckMenu_CodeMetrics";
        public override int DisplayOrder => (int)NavigationMenuItemDisplayOrder.CodeMetrics;
    }
}
