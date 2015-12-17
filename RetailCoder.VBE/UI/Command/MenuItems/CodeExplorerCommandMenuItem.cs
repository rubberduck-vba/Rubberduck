using System.Windows.Input;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class CodeExplorerCommandMenuItem : CommandMenuItemBase
    {
        public CodeExplorerCommandMenuItem(ICommand command) 
            : base(command)
        {
        }

        public override string Key { get { return "RubberduckMenu_CodeExplorer"; } }
        public override int DisplayOrder { get { return (int)NavigationMenuItemDisplayOrder.CodeExplorer; } }

        public override bool EvaluateCanExecute(IRubberduckParserState state)
        {
            return state.Status == ParserState.Ready ||
                   state.Status == ParserState.Resolving;
        }
    }
}