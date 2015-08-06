using System.Collections.Generic;
using Microsoft.Office.Core;

namespace Rubberduck.UI.Command
{
    public class UnitTestingParentMenu : ParentMenuItemBase
    {
        public UnitTestingParentMenu(CommandBarControls parent, IEnumerable<IMenuItem> items) 
            : base(parent, RubberduckUI.RubberduckMenu_UnitTests, items)
        {
        }

        public override int DisplayOrder { get { return (int)RubberduckMenuItemDisplayOrder.UnitTesting; } }
    }
}