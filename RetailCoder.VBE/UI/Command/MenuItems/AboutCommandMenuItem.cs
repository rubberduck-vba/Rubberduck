using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class AboutCommandMenuItem : CommandMenuItemBase
    {
        public AboutCommandMenuItem(CommandBase command) : base(command)
        {
        }

        public override string Key { get { return "RubberduckMenu_About"; } }
        public override bool BeginGroup { get { return true; } }
        public override int DisplayOrder { get { return (int)RubberduckMenuItemDisplayOrder.About; } }

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return true;
        }
    }
}
