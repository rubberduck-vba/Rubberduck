using System.Drawing;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems.ParentMenus;
using Rubberduck.UI.Command.Refactorings;

namespace Rubberduck.UI.Command.MenuItems
{
    public class RefactorImplementInterfaceCommandMenuItem : CommandMenuItemBase
    {
        public RefactorImplementInterfaceCommandMenuItem(RefactorImplementInterfaceCommand command) 
            : base(command)
        {
        }

        public override string Key => "RefactorMenu_ImplementInterface";
        public override int DisplayOrder => (int)RefactoringsMenuItemDisplayOrder.ImplementInterface;
        public override Image Image => Resources.CommandBarIcons.ImplementInterface;
        public override Image Mask => Resources.CommandBarIcons.ImplementInterfaceMask;

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && Command.CanExecute(null);
        }
    }
}
