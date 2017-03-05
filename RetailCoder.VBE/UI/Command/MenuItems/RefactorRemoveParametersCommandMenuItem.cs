using System.Drawing;
using Rubberduck.Parsing.VBA;
using Rubberduck.Properties;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class RefactorRemoveParametersCommandMenuItem : CommandMenuItemBase
    {
        public RefactorRemoveParametersCommandMenuItem(CommandBase command) : base(command)
        {
        }

        public override string Key { get { return "RefactorMenu_RemoveParameter"; } }
        public override int DisplayOrder { get { return (int)RefactoringsMenuItemDisplayOrder.RemoveParameters; } }
        public override bool BeginGroup { get { return true; } }

        public override Image Image { get { return Resources.RemoveParameters; } }
        public override Image Mask { get { return Resources.RemoveParametersMask; }}

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && Command.CanExecute(null);
        }
    }
}
