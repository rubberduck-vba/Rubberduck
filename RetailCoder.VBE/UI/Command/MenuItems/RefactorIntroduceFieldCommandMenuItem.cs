using System.Windows.Input;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class RefactorIntroduceFieldCommandMenuItem : CommandMenuItemBase
    {
        public RefactorIntroduceFieldCommandMenuItem (ICommand command)
            : base(command)
        {
        }

        public override string Key { get { return "RefactorMenu_IntroduceField"; } }
        public override int DisplayOrder { get { return (int)RefactoringsMenuItemDisplayOrder.IntroduceField; } }

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state.Status == ParserState.Ready;
        }
    }
}