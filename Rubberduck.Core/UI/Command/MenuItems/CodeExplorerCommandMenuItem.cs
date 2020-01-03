using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.ComCommands;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class CodeExplorerCommandMenuItem : CommandMenuItemBase
    {
        public CodeExplorerCommandMenuItem(CodeExplorerCommand command) 
            : base(command)
        {
        }

        public override bool EvaluateCanExecute(RubberduckParserState state) => true;

        public override string Key => "RubberduckMenu_CodeExplorer";
        public override int DisplayOrder => (int)NavigationMenuItemDisplayOrder.CodeExplorer;
    }
}
