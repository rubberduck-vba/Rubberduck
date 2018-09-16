using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class SettingsCommandMenuItem : CommandMenuItemBase
    {
        public SettingsCommandMenuItem(SettingsCommand command) : base(command)
        {
        }

        public override string Key => "RubberduckMenu_Settings";
        public override bool BeginGroup => true;
        public override int DisplayOrder => (int)RubberduckMenuItemDisplayOrder.Settings;

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return true;
        }
    }
}
