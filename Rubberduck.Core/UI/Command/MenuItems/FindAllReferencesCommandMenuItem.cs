using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.ComCommands;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public abstract class FindAllReferencesCommandMenuItemBase : CommandMenuItemBase
    {
        protected FindAllReferencesCommandMenuItemBase(FindAllReferencesCommand command)
            : base(command)
        {}

        public override string Key => "ContextMenu_FindAllReferences";

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && Command.CanExecute(null);
        }
    }


    public class FindAllReferencesCommandMenuItem : FindAllReferencesCommandMenuItemBase
    {
        public FindAllReferencesCommandMenuItem(FindAllReferencesCommand command)
            : base(command)
        {}
        
        public override int DisplayOrder => (int)CodePaneContextMenuItemDisplayOrder.FindAllReferences;
    }


    public class ProjectExplorerFindAllReferencesCommandMenuItem : FindAllReferencesCommandMenuItemBase
    {
        public ProjectExplorerFindAllReferencesCommandMenuItem(ProjectExplorerFindAllReferencesCommand command)
            : base(command)
        {}

        public override int DisplayOrder => (int)ProjectExplorerContextMenuItemDisplayOrder.FindAllReferences;
    }
}
