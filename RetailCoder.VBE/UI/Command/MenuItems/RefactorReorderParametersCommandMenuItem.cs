using System.Drawing;
using System.Windows.Input;
using Rubberduck.Parsing.VBA;
using Rubberduck.Properties;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class RefactorReorderParametersCommandMenuItem : CommandMenuItemBase
    {
        public RefactorReorderParametersCommandMenuItem(ICommand command) : base(command)
        {
        }

        public override string Key { get { return "RefactorMenu_ReorderParameters"; } }
        public override int DisplayOrder { get { return (int)RefactoringsMenuItemDisplayOrder.ReorderParameters; } }
        public override Image Image { get { return Resources.ReorderParameters_6780_32; } }
        public override Image Mask { get { return Resources.ReorderParameters_6780_32_Mask; } }

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state.Status == ParserState.Ready;
        }
    }
}