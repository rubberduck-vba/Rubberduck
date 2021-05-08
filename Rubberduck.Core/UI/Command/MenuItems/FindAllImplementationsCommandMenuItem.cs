using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.ComCommands;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public abstract class FindAllImplementationsCommandMenuItemBase : CommandMenuItemBase
    {
        protected FindAllImplementationsCommandMenuItemBase(FindAllImplementationsCommand command) 
            : base(command)
        {}

        public override string Key => "ContextMenu_GoToImplementation";

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && Command.CanExecute(null);
        }
    }

    public class FindAllImplementationsCommandMenuItem : FindAllImplementationsCommandMenuItemBase
    {
        public FindAllImplementationsCommandMenuItem(FindAllImplementationsCommand command) 
            : base(command)
        {}

        public override int DisplayOrder => (int)CodePaneContextMenuItemDisplayOrder.FindAllImplementations;
    }

    public class ProjectExplorerFindAllImplementationsCommandMenuItem : FindAllImplementationsCommandMenuItemBase
    {
        public ProjectExplorerFindAllImplementationsCommandMenuItem(ProjectExplorerFindAllImplementationsCommand command)
            : base(command)
        {}

        public override int DisplayOrder => (int)ProjectExplorerContextMenuItemDisplayOrder.FindAllImplementations;
    }
}
