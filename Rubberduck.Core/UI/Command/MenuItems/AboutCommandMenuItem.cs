using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class AboutCommandMenuItem : CommandMenuItemBase
    {
        public AboutCommandMenuItem(AboutCommand command) : base(command)
        {
        }

        public override string Key => "RubberduckMenu_About";
        public override bool BeginGroup => true;
        public override int DisplayOrder => (int)RubberduckMenuItemDisplayOrder.About;

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return true;
        }
    }
}
