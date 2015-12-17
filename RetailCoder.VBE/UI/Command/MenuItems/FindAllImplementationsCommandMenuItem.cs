using System.Windows.Input;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class FindAllImplementationsCommandMenuItem : CommandMenuItemBase
    {
        public FindAllImplementationsCommandMenuItem(ICommand command) : base(command)
        {
        }

        public override string Key { get { return "ContextMenu_GoToImplementation"; } }
        public override int DisplayOrder { get { return (int)NavigationMenuItemDisplayOrder.FindImplementations; } }

        public override bool EvaluateCanExecute(IRubberduckParserState state)
        {
            return state.Status == ParserState.Ready;
        }
    }
}