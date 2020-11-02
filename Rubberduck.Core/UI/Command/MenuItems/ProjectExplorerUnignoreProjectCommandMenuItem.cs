using Rubberduck.UI.Command.ComCommands;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class ProjectExplorerUnignoreProjectCommandMenuItem : CommandMenuItemBase
    {
        public ProjectExplorerUnignoreProjectCommandMenuItem(ProjectExplorerUnignoreProjectCommand command)
            : base(command)
        { }

        public override string Key => "ProjectExplorer_UnignoreProject";
        public override int DisplayOrder => (int)ProjectExplorerContextMenuItemDisplayOrder.UnignoreProject;
    }
}