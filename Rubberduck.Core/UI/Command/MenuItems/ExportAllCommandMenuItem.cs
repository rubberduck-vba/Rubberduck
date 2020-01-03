using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.ComCommands;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class ExportAllCommandMenuItem : CommandMenuItemBase
    {
        public ExportAllCommandMenuItem(ExportAllCommand command) : base(command)
        {
        }

        public override string Key => "ToolsMenu_ExportProject";

        public override int DisplayOrder => (int)ToolsMenuItemDisplayOrder.ExportAll;

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return Command.CanExecute(null);
        }
    }
}
