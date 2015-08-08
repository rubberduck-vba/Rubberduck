using System.Collections.Generic;

namespace Rubberduck.UI.Command
{
    public class UnitTestingParentMenu : ParentMenuItemBase
    {
        public UnitTestingParentMenu(IEnumerable<IMenuItem> items) 
            : base("RubberduckMenu_UnitTests", items)
        {
        }

        public override int DisplayOrder { get { return (int)RubberduckMenuItemDisplayOrder.UnitTesting; } }
    }
}