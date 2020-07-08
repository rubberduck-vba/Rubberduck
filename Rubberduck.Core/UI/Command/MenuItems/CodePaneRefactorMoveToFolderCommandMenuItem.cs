using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems.ParentMenus;
using Rubberduck.UI.Command.Refactorings;

namespace Rubberduck.UI.Command.MenuItems
{
    public class CodePaneRefactorMoveToFolderCommandMenuItem : CommandMenuItemBase
    {
        public CodePaneRefactorMoveToFolderCommandMenuItem(CodePaneRefactorMoveToFolderCommand command)
            : base(command)
        {}

        public override string Key => "RefactorMenu_MoveToFolder";
        public override int DisplayOrder => (int)RefactoringsMenuItemDisplayOrder.MoveToFolder;
        public override bool BeginGroup => true;

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && Command.CanExecute(null);
        }
    }
}
