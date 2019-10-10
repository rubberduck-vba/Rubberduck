using System.Drawing;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.ComCommands;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class FindSymbolCommandMenuItem : CommandMenuItemBase
    {
        public FindSymbolCommandMenuItem(FindSymbolCommand command) 
            : base(command)
        {
        }

        public override string Key => "ContextMenu_FindSymbol";
        public override int DisplayOrder => (int)CodePaneContextMenuItemDisplayOrder.FindSymbol;
        public override bool BeginGroup => true;

        public override Image Image => Resources.CommandBarIcons.FindSymbol;
        public override Image Mask => Resources.CommandBarIcons.FindSymbolMask;

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && state.Status >= ParserState.ResolvedDeclarations && state.Status < ParserState.Error;
        }
    }
}
