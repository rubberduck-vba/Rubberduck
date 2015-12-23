using System.Windows.Input;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class RefactorExtractInterfaceCommandMenuItem : CommandMenuItemBase
    {
        public RefactorExtractInterfaceCommandMenuItem(ICommand command) 
            : base(command)
        {
        }

        public override string Key { get { return "RefactorMenu_ExtractInterface"; } }
        public override int DisplayOrder { get { return (int)RefactoringsMenuItemDisplayOrder.ExtractInterface; } }

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state.Status == ParserState.Ready;
        }
    }
}