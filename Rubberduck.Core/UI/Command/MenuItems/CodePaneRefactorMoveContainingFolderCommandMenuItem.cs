using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.ComCommands;
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

    public class PeekDefinitionCommandMenuItem : CommandMenuItemBase
    {
        public PeekDefinitionCommandMenuItem(PeekDefinitionCommand command)
            : base(command)
        {}

        public override string Key => "ContextMenu_PeekDefinition";
        public override bool BeginGroup => true;

        public override int DisplayOrder => (int) CodePaneContextMenuItemDisplayOrder.PeekDefinition;

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && Command.CanExecute(null);
        }
    }
}
