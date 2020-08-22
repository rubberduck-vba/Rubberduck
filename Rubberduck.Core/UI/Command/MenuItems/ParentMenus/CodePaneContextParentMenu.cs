using System.Collections.Generic;

namespace Rubberduck.UI.Command.MenuItems.ParentMenus
{
    public class CodePaneContextParentMenu : ParentMenuItemBase
    {
        public CodePaneContextParentMenu(IEnumerable<IMenuItem> items, int beforeIndex)
            : base("RubberduckMenu", items, beforeIndex)
        {
        }

        public override bool BeginGroup => true;
    }

    public enum CodePaneContextMenuItemDisplayOrder
    {
        Refactorings,
        Annotate,
        Indenter,
        RegexSearchReplace,
        FindSymbol,
        FindAllReferences,
        FindAllImplementations,
        RunSelectedTestModule,
        RunSelectedTest
    }
}
