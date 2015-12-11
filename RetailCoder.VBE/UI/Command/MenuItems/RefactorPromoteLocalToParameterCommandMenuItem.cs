using System.Windows.Input;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class RefactorPromoteLocalToParameterCommandMenuItem : CommandMenuItemBase
    {
        public RefactorPromoteLocalToParameterCommandMenuItem (ICommand command)
            : base(command)
        {
        }

        public override string Key { get { return "RefactorMenu_PromoteLocalToParameter"; } }
        public override int DisplayOrder { get { return (int)RefactoringsMenuItemDisplayOrder.PromoteLocalToParameter; } }

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state.Status == ParserState.Ready;
        }
    }
}