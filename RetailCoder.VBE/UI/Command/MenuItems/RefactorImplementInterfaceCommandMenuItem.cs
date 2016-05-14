using System.Drawing;
using System.Windows.Input;
using Rubberduck.Parsing.VBA;
using Rubberduck.Properties;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class RefactorImplementInterfaceCommandMenuItem : CommandMenuItemBase
    {
        public RefactorImplementInterfaceCommandMenuItem(ICommand command) 
            : base(command)
        {
        }

        public override string Key { get { return "RefactorMenu_ImplementInterface"; } }
        public override int DisplayOrder { get { return (int)RefactoringsMenuItemDisplayOrder.ImplementInterface; } }
        public override Image Image { get { return Resources.ImplementInterface_5540_32; } }
        public override Image Mask { get { return Resources.ImplementInterface_5540_321_Mask; } }

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return Command.CanExecute(null);
        }
    }
}
