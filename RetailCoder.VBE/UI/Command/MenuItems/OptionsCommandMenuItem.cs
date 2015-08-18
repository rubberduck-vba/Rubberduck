using System.Windows.Input;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class OptionsCommandMenuItem : CommandMenuItemBase
    {
        public OptionsCommandMenuItem(ICommand command) : base(command)
        {
        }

        public override string Key { get { return "RubberduckMenu_Options"; } }
        public override bool BeginGroup { get { return true; } }
        public override int DisplayOrder { get { return (int)RubberduckMenuItemDisplayOrder.Options; } }
    }
}