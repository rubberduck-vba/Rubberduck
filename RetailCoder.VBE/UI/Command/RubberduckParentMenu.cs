using System.Collections.Generic;
using Microsoft.Office.Core;

namespace Rubberduck.UI.Command
{
    public class RubberduckParentMenu : ParentMenuItemBase
    {
        public RubberduckParentMenu(CommandBarControls parent, IEnumerable<IMenuItem> items, int beforeIndex) 
            : base(parent, RubberduckUI.RubberduckMenu, items, beforeIndex)
        {
        }
    }
}
