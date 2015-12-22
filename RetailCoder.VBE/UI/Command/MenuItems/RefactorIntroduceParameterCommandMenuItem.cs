using System.Windows.Input;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class RefactorIntroduceParameterCommandMenuItem : CommandMenuItemBase
    {
        public RefactorIntroduceParameterCommandMenuItem (ICommand command)
            : base(command)
        {
        }

        public override string Key { get { return "RefactorMenu_IntroduceParameter"; } }
        public override int DisplayOrder { get { return (int)RefactoringsMenuItemDisplayOrder.IntroduceParameter; } }

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state.Status == ParserState.Ready;
        }
    }
}