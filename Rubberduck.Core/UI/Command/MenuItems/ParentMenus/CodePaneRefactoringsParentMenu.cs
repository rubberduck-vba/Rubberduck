using System.Collections.Generic;
using Rubberduck.Parsing.UIContext;

namespace Rubberduck.UI.Command.MenuItems.ParentMenus
{
    public class CodePaneRefactoringsParentMenu : ParentMenuItemBase
    {
        public CodePaneRefactoringsParentMenu(IEnumerable<IMenuItem> items, IUiDispatcher dispatcher)
            : base(dispatcher, "RubberduckMenu_CodePaneRefactor", items)
        { }

        //This display order is different from the main menu; it is the sole reason this class is separate from the main menu one.
        public override int DisplayOrder => (int)CodePaneContextMenuItemDisplayOrder.Refactorings;
    }
}
