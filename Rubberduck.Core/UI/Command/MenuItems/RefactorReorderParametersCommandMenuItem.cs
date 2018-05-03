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

        public override string Key => "RefactorMenu_ReorderParameters";
        public override int DisplayOrder => (int)RefactoringsMenuItemDisplayOrder.ReorderParameters;
        public override Image Image => Resources.ReorderParameters;
        public override Image Mask => Resources.ReorderParametersMask;

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && Command.CanExecute(null);
        }
    }
}
