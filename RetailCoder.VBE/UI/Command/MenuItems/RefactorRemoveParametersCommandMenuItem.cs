using System.Drawing;
using System.Windows.Input;
using Rubberduck.Parsing.VBA;
using Rubberduck.Properties;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class RefactorRemoveParametersCommandMenuItem : CommandMenuItemBase
    {
        public RefactorRemoveParametersCommandMenuItem(ICommand command) : base(command)
        {
        }

        public override string Key { get { return "RefactorMenu_RemoveParameter"; } }
        public override int DisplayOrder { get { return (int)RefactoringsMenuItemDisplayOrder.RemoveParameters; } }
        public override Image Image { get { return Resources.RemoveParameters_6781_32; } }
        public override Image Mask { get { return Resources.RemoveParameters_6781_32_Mask; }}

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state.Status == RubberduckParserState.State.Ready;
        }
    }
}