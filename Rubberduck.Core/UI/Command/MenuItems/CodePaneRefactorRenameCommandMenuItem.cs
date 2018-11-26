using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems.ParentMenus;
using Rubberduck.UI.Command.Refactorings;

namespace Rubberduck.UI.Command.MenuItems
{
    public class CodePaneRefactorRenameCommandMenuItem : CommandMenuItemBase
    {
        public CodePaneRefactorRenameCommandMenuItem(CodePaneRefactorRenameCommand command)
            : base(command)
        {
        }

        public override string Key => "RefactorMenu_Rename";
        public override int DisplayOrder => (int)RefactoringsMenuItemDisplayOrder.RenameIdentifier;

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && state.Status == ParserState.Ready && Command.CanExecute(null);
        }
    }
}
