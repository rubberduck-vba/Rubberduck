using System.Collections.Generic;

namespace Rubberduck.UI.Command
{
    public class RefactoringsParentMenu : ParentMenuItemBase
    {
        public RefactoringsParentMenu(IEnumerable<IMenuItem> items)
            : base("RubberduckMenu_Refactor", items)
        {
        }

        public override int DisplayOrder { get { return (int)RubberduckMenuItemDisplayOrder.Refactorings; } }
    }
}