using System.Drawing;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems.ParentMenus;
using Rubberduck.UI.Command.Refactorings;

namespace Rubberduck.UI.Command.MenuItems
{
    public class RefactorRemoveParametersCommandMenuItem : CommandMenuItemBase
    {
        public RefactorRemoveParametersCommandMenuItem(RefactorRemoveParametersCommand command) : base(command)
        {
        }

        public override string Key => "RefactorMenu_RemoveParameter";
        public override int DisplayOrder => (int)RefactoringsMenuItemDisplayOrder.RemoveParameters;
        public override bool BeginGroup => true;

        public override Image Image => Resources.CommandBarIcons.RemoveParameters;
        public override Image Mask => Resources.CommandBarIcons.RemoveParametersMask;

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && Command.CanExecute(null);
        }
    }
}
