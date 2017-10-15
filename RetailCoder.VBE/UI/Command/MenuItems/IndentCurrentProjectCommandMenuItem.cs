using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class IndentCurrentProjectCommandMenuItem : CommandMenuItemBase
    {
        public IndentCurrentProjectCommandMenuItem(CommandBase command) : base(command) { }

        public override string Key { get { return "IndentCurrentProject"; } }
        public override int DisplayOrder { get { return (int)SmartIndenterMenuItemDisplayOrder.CurrentProject; } }
    }
}
