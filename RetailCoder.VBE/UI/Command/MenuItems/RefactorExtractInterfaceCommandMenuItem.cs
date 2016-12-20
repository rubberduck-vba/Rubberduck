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

        public override string Key { get { return "RefactorMenu_ExtractInterface"; } }
        public override int DisplayOrder { get { return (int)RefactoringsMenuItemDisplayOrder.ExtractInterface; } }
        public override Image Image { get { return Resources.ExtractInterface; } }
        public override Image Mask { get { return Resources.ExtractInterfaceMask; } }

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && Command.CanExecute(null);
        }
    }
}
