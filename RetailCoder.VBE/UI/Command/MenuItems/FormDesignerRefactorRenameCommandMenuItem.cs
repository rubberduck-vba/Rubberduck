using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class FormDesignerRefactorRenameCommandMenuItem : CommandMenuItemBase
    {
        public FormDesignerRefactorRenameCommandMenuItem(CommandBase command)
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
