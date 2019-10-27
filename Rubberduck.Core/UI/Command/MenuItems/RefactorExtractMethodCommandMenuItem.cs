using System.Drawing;
using Rubberduck.Parsing.Common;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems.ParentMenus;
using Rubberduck.UI.Command.Refactorings;

namespace Rubberduck.UI.Command.MenuItems
{
    [Disabled]
    public class RefactorExtractMethodCommandMenuItem : CommandMenuItemBase
    {
        public RefactorExtractMethodCommandMenuItem(RefactorExtractMethodCommand command) 
            : base(command)
        {
        }

        public override string Key => "RefactorMenu_ExtractMethod";
        public override int DisplayOrder => (int)RefactoringsMenuItemDisplayOrder.ExtractMethod;

        public override bool BeginGroup => true;

        public override Image Image => Resources.CommandBarIcons.ExtractMethod;
        public override Image Mask => Resources.CommandBarIcons.ExtractMethodMask;

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && Command.CanExecute(null);
        }
    }
}
