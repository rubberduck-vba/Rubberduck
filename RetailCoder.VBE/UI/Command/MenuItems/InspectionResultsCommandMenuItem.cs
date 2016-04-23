using System.Windows.Input;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class InspectionResultsCommandMenuItem : CommandMenuItemBase
    {
        public InspectionResultsCommandMenuItem(ICommand command) 
            : base(command)
        {
        }

        public override string Key { get { return "RubberduckMenu_CodeInspections"; } }
        public override int DisplayOrder { get { return (int)RubberduckMenuItemDisplayOrder.CodeInspections; } }

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state.Status == ParserState.Ready;
        }
    }
}