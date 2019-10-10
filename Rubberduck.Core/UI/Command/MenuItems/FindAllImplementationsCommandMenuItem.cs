using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.ComCommands;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class FindAllImplementationsCommandMenuItem : CommandMenuItemBase
    {
        public FindAllImplementationsCommandMenuItem(FindAllImplementationsCommand command) : base(command)
        {
        }

        public override string Key => "ContextMenu_GoToImplementation";
        public override int DisplayOrder => (int)CodePaneContextMenuItemDisplayOrder.FindAllImplementations;

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && Command.CanExecute(null);
        }
    }
}
