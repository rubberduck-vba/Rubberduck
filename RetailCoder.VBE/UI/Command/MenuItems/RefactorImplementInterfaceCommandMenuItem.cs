using System.Drawing;
using Rubberduck.Parsing.VBA;
using Rubberduck.Properties;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class RefactorImplementInterfaceCommandMenuItem : CommandMenuItemBase
    {
        public RefactorImplementInterfaceCommandMenuItem(CommandBase command) 
            : base(command)
        {
        }

        public override string Key { get { return "RefactorMenu_ImplementInterface"; } }
        public override int DisplayOrder { get { return (int)RefactoringsMenuItemDisplayOrder.ImplementInterface; } }
        public override Image Image { get { return Resources.ImplementInterface; } }
        public override Image Mask { get { return Resources.ImplementInterfaceMask; } }

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && Command.CanExecute(null);
        }
    }
}
