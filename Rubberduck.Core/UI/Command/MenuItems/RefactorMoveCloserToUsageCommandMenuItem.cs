using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems.ParentMenus;
using Rubberduck.UI.Command.Refactorings;

namespace Rubberduck.UI.Command.MenuItems
{
    public class RefactorMoveCloserToUsageCommandMenuItem : CommandMenuItemBase
    {
        public RefactorMoveCloserToUsageCommandMenuItem(RefactorMoveCloserToUsageCommand command)
            : base(command)
        {
        }

        public override string Key => "RefactorMenu_MoveCloserToUsage";
        public override int DisplayOrder => (int)RefactoringsMenuItemDisplayOrder.MoveCloserToUsage;
        public override bool BeginGroup => true;

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && Command.CanExecute(null);
        }
    }
}
