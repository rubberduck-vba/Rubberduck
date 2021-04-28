using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems.ParentMenus;
using Rubberduck.UI.Command.Refactorings;

namespace Rubberduck.UI.Command.MenuItems
{
    public class CodePaneRefactorMoveContainingFolderCommandMenuItem : CommandMenuItemBase
    {
        public CodePaneRefactorMoveContainingFolderCommandMenuItem(CodePaneRefactorMoveContainingFolderCommand command)
            : base(command)
        {}

        public override string Key => "RefactorMenu_MoveContainingFolder";
        public override int DisplayOrder => (int)RefactoringsMenuItemDisplayOrder.MoveContainingFolder;
        public override bool BeginGroup => false;

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && Command.CanExecute(null);
        }
    }
}
