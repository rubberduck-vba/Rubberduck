using System.Drawing;
using System.Windows.Input;
using Rubberduck.Parsing.VBA;
using Rubberduck.Properties;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class RefactorExtractInterfaceCommandMenuItem : CommandMenuItemBase
    {
        public RefactorExtractInterfaceCommandMenuItem(ICommand command) 
            : base(command)
        {
        }

        public override string Key { get { return "RefactorMenu_ExtractInterface"; } }
        public override int DisplayOrder { get { return (int)RefactoringsMenuItemDisplayOrder.ExtractInterface; } }
        public override Image Image { get { return Resources.ExtractInterface_6778_32; } }
        public override Image Mask { get { return Resources.ExtractInterface_6778_321_Mask; } }

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return Command.CanExecute(null);
        }
    }
}