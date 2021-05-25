using System.Collections.Generic;
using Rubberduck.Parsing.UIContext;

namespace Rubberduck.UI.Command.MenuItems.ParentMenus
{
    public class ProjectWindowContextParentMenu : ParentMenuItemBase
    {
        public ProjectWindowContextParentMenu(IEnumerable<IMenuItem> items, int beforeIndex, IUiDispatcher dispatcher)
            : base(dispatcher,"RubberduckMenu", items, beforeIndex)
        {
        }

        public override bool BeginGroup => true;
    }

    public enum ProjectExplorerContextMenuItemDisplayOrder
    {
        RenameIdentifier,
        PeekDefinition,
        FindSymbol,
        FindAllReferences,
        FindAllImplementations,
        AddRemoveReferences,
        IgnoreProject,
        UnignoreProject
    }
}
