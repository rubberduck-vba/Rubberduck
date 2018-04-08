using System.Drawing;
using Rubberduck.Parsing.VBA;
using Rubberduck.Properties;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class RefactorExtractInterfaceCommandMenuItem : CommandMenuItemBase
    {
        public RefactorExtractInterfaceCommandMenuItem(CommandBase command) 
            : base(command)
        {
        }

        public override string Key => "RefactorMenu_ExtractInterface";
        public override int DisplayOrder => (int)RefactoringsMenuItemDisplayOrder.ExtractInterface;
        public override Image Image => Resources.ExtractInterface;
        public override Image Mask => Resources.ExtractInterfaceMask;

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && Command.CanExecute(null);
        }
    }
}
