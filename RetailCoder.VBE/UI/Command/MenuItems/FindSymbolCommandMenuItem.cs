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

        public override string Key { get {return "ContextMenu_FindSymbol"; } }
        public override int DisplayOrder { get { return (int)CodePaneContextMenuItemDisplayOrder.FindSymbol; } }
        public override bool BeginGroup { get { return true; } }

        public override Image Image { get { return Resources.FindSymbol; } }
        public override Image Mask { get { return Resources.FindSymbolMask; } }

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && state.Status >= ParserState.ResolvedDeclarations && state.Status < ParserState.Error;
        }
    }
}
