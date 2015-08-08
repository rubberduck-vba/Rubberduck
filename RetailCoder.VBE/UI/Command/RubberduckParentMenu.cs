using System.Collections.Generic;

namespace Rubberduck.UI.Command
{
    public class RubberduckParentMenu : ParentMenuItemBase
    {
        public RubberduckParentMenu(IEnumerable<IMenuItem> items, int beforeIndex) 
            : base("RubberduckMenu", items, beforeIndex)
        {
        }
    }
}
