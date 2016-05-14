using System.Windows.Input;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class RefactorMoveCloserToUsageCommandMenuItem : CommandMenuItemBase
    {
        public RefactorMoveCloserToUsageCommandMenuItem(ICommand command)
            : base(command)
        {
        }

        public override string Key { get { return "RefactorMenu_MoveCloserToUsage"; } }
        public override int DisplayOrder { get { return (int)RefactoringsMenuItemDisplayOrder.MoveCloserToUsage; } }
        public override bool BeginGroup { get { return true; } }

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return Command.CanExecute(null);
        }
    }
}
