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

        public override string Key { get { return "RefactorMenu_Rename"; } }
        public override int DisplayOrder { get { return (int)RefactoringsMenuItemDisplayOrder.RenameIdentifier; } }

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && state.Status == ParserState.Ready;
        }
    }
}
