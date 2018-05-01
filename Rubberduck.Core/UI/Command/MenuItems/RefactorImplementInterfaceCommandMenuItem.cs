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

        public override string Key => "RefactorMenu_ImplementInterface";
        public override int DisplayOrder => (int)RefactoringsMenuItemDisplayOrder.ImplementInterface;
        public override Image Image => Resources.ImplementInterface;
        public override Image Mask => Resources.ImplementInterfaceMask;
        public override byte[] LowColorImageBytes => Resources.ImplementInterface_LowColor;

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && Command.CanExecute(null);
        }
    }
}
