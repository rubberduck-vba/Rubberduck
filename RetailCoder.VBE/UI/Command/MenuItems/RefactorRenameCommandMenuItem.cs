using System;
using System.Windows.Input;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class RefactorRenameCommandMenuItem : CommandMenuItemBase
    {
        public RefactorRenameCommandMenuItem(ICommand command)
            : base(command)
        {
        }

        public override string Key { get { return "RefactorMenu_Rename"; } }
        public override int DisplayOrder { get { return (int)RefactoringsMenuItemDisplayOrder.RenameIdentifier; } }

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state.Status == ParserState.Ready;
        }
    }
}