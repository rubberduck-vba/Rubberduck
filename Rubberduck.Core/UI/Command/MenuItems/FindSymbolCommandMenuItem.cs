using System.Drawing;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.ComCommands;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public abstract class FindSymbolCommandMenuItemBase : CommandMenuItemBase
    {
        protected FindSymbolCommandMenuItemBase(FindSymbolCommand command)
            : base(command)
        {}

        public override string Key => "ContextMenu_FindSymbol";
        public override Image Image => Resources.CommandBarIcons.FindSymbol;
        public override Image Mask => Resources.CommandBarIcons.FindSymbolMask;

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && state.Status >= ParserState.ResolvedDeclarations && state.Status < ParserState.Error;
        }
    }

    public class FindSymbolCommandMenuItem : FindSymbolCommandMenuItemBase
    {
        public FindSymbolCommandMenuItem(FindSymbolCommand command) 
            : base(command)
        {}

        public override int DisplayOrder => (int)CodePaneContextMenuItemDisplayOrder.FindSymbol;
    }

    public class ProjectExplorerFindSymbolCommandMenuItem : FindSymbolCommandMenuItemBase
    {
        public ProjectExplorerFindSymbolCommandMenuItem(FindSymbolCommand command)
            : base(command)
        {}

        public override int DisplayOrder => (int)ProjectExplorerContextMenuItemDisplayOrder.FindSymbol;
    }
}
