using System.Drawing;
using Rubberduck.Parsing.VBA;
using Rubberduck.Properties;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class FindSymbolCommandMenuItem : CommandMenuItemBase
    {
        public FindSymbolCommandMenuItem(CommandBase command) 
            : base(command)
        {
        }

        public override string Key => "ContextMenu_FindSymbol";
        public override int DisplayOrder => (int)CodePaneContextMenuItemDisplayOrder.FindSymbol;
        public override bool BeginGroup => true;

        public override Image Image => Resources.FindSymbol;
        public override Image Mask => Resources.FindSymbolMask;

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && state.Status >= ParserState.ResolvedDeclarations && state.Status < ParserState.Error;
        }
    }
}
