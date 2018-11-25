using System.Drawing;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems.ParentMenus;
using Rubberduck.UI.Command.Refactorings;

namespace Rubberduck.UI.Command.MenuItems
{
    public class RefactorReorderParametersCommandMenuItem : CommandMenuItemBase
    {
        public RefactorReorderParametersCommandMenuItem(RefactorReorderParametersCommand command) : base(command)
        {
        }

        public override string Key => "RefactorMenu_ReorderParameters";
        public override int DisplayOrder => (int)RefactoringsMenuItemDisplayOrder.ReorderParameters;
        public override Image Image => Resources.CommandBarIcons.ReorderParameters;
        public override Image Mask => Resources.CommandBarIcons.ReorderParametersMask;

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && Command.CanExecute(null);
        }
    }
}
