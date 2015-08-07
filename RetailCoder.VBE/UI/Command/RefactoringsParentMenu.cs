using System.Collections.Generic;
using Microsoft.Office.Core;

namespace Rubberduck.UI.Command
{
    public class RefactoringsParentMenu : ParentMenuItemBase
    {
        public RefactoringsParentMenu(CommandBarControls parent, IEnumerable<IMenuItem> items)
            : base(parent, RubberduckUI.RubberduckMenu_Refactor, items)
        {
        }

        public override int DisplayOrder { get { return (int)RubberduckMenuItemDisplayOrder.Refactorings; } }
    }
}