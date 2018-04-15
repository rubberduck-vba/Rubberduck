using System.Collections.Generic;

namespace Rubberduck.UI.Command.MenuItems.ParentMenus
{
    public class CodePaneContextParentMenu : ParentMenuItemBase
    {
        public CodePaneContextParentMenu(ICommandBarButtonFactory buttonFactory, IEnumerable<IMenuItem> items, int beforeIndex)
            : base(buttonFactory, "RubberduckMenu", items, beforeIndex)
        {
        }

        public override bool BeginGroup => true;
    }

    public enum CodePaneContextMenuItemDisplayOrder
    {
        Refactorings,
        Indenter,
        RegexSearchReplace,
        FindSymbol,
        FindAllReferences,
        FindAllImplementations,
    }
}
