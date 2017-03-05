using System.Drawing;
using Rubberduck.Parsing.VBA;
using Rubberduck.Properties;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class RefactorReorderParametersCommandMenuItem : CommandMenuItemBase
    {
        public RefactorReorderParametersCommandMenuItem(CommandBase command) : base(command)
        {
        }

        public override string Key { get { return "RefactorMenu_ReorderParameters"; } }
        public override int DisplayOrder { get { return (int)RefactoringsMenuItemDisplayOrder.ReorderParameters; } }
        public override Image Image { get { return Resources.ReorderParameters; } }
        public override Image Mask { get { return Resources.ReorderParametersMask; } }

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && Command.CanExecute(null);
        }
    }
}
