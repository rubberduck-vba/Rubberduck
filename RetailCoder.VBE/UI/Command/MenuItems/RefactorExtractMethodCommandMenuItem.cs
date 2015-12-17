using System.Drawing;
using System.Windows.Input;
using Rubberduck.Parsing.VBA;
using Rubberduck.Properties;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class RefactorExtractMethodCommandMenuItem : CommandMenuItemBase
    {
        public RefactorExtractMethodCommandMenuItem(ICommand command) 
            : base(command)
        {
        }

        public override string Key { get { return "RefactorMenu_ExtractMethod"; } }
        public override int DisplayOrder { get { return (int)RefactoringsMenuItemDisplayOrder.ExtractMethod; } }
        public override Image Image { get { return Resources.ExtractMethod_6786_32; } }
        public override Image Mask { get { return Resources.ExtractMethod_6786_32_Mask; } }

        public override bool EvaluateCanExecute(IRubberduckParserState state)
        {
            return state.Status == ParserState.Ready;
        }
    }
}