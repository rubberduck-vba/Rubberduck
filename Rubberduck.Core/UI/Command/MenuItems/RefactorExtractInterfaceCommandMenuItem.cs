using System.Drawing;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems.ParentMenus;
using Rubberduck.UI.Command.Refactorings;

namespace Rubberduck.UI.Command.MenuItems
{
    public class RefactorExtractInterfaceCommandMenuItem : CommandMenuItemBase
    {
        public RefactorExtractInterfaceCommandMenuItem(RefactorExtractInterfaceCommand command) 
            : base(command)
        {
        }

        public override string Key => "RefactorMenu_ExtractInterface";
        public override int DisplayOrder => (int)RefactoringsMenuItemDisplayOrder.ExtractInterface;
        public override Image Image => Resources.CommandBarIcons.ExtractInterface;
        public override Image Mask => Resources.CommandBarIcons.ExtractInterfaceMask;

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && Command.CanExecute(null);
        }
    }
}
