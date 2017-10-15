using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class SettingsCommandMenuItem : CommandMenuItemBase
    {
        public SettingsCommandMenuItem(CommandBase command) : base(command)
        {
        }

        public override string Key { get { return "RubberduckMenu_Settings"; } }
        public override bool BeginGroup { get { return true; } }
        public override int DisplayOrder { get { return (int)RubberduckMenuItemDisplayOrder.Settings; } }

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return true;
        }
    }
}
