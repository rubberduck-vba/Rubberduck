using Rubberduck.UI.Command.MenuItems.ParentMenus;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace Rubberduck.UI.Command.MenuItems
{
    class RegexAssistantCommandMenuItem : CommandMenuItemBase
    {
        public RegexAssistantCommandMenuItem(ICommand command) : base(command)
        {
        }

        public override string Key { get { return "RubberduckMenu_RegexAssistant"; } }

        public override int DisplayOrder { get { return (int)RubberduckMenuItemDisplayOrder.RegexAssistant; } }
    }
}
