using System.Collections.Generic;

namespace Rubberduck.UI.Command.MenuItems.ParentMenus
{
    public class CodePaneContextParentMenu : ParentMenuItemBase
    {
        public CodePaneContextParentMenu(IEnumerable<IMenuItem> items, int beforeIndex)
            : base("RubberduckMenu", items, beforeIndex)
        {
        }

        public override bool BeginGroup { get { return true; } }

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
