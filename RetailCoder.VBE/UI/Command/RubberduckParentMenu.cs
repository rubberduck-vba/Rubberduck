using System.Collections.Generic;
using Microsoft.Office.Core;

namespace Rubberduck.UI.Command
{
    public class RubberduckParentMenu : ParentMenuItemBase
    {
        public RubberduckParentMenu(CommandBarControls parent, int beforeIndex, IEnumerable<IMenuItem> items) 
            : base(parent, RubberduckUI.RubberduckMenu, items, beforeIndex)
        {
        }
    }
}
