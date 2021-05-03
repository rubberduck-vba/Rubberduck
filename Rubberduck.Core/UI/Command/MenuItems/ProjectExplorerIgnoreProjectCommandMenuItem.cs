using Rubberduck.UI.Command.ComCommands;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class ProjectExplorerIgnoreProjectCommandMenuItem : CommandMenuItemBase
    {
        public ProjectExplorerIgnoreProjectCommandMenuItem(ProjectExplorerIgnoreProjectCommand command) 
            : base(command)
        { }

        public override string Key => "ProjectExplorer_IgnoreProject";
        public override int DisplayOrder => (int)ProjectExplorerContextMenuItemDisplayOrder.IgnoreProject;
        public override bool BeginGroup => true;
    }
}