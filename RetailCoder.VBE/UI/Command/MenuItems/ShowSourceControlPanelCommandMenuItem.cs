using System.Windows.Input;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class ShowSourceControlPanelCommandMenuItem : CommandMenuItemBase
    {
        public ShowSourceControlPanelCommandMenuItem(ICommand command) 
            : base(command)
        {
        }

        public override string Key { get { return "RubberduckMenu_SourceControl"; } }
        public override int DisplayOrder { get { return (int)RubberduckMenuItemDisplayOrder.SourceControl; } }
    }
}