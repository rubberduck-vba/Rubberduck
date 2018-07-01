using System.Drawing;
using Rubberduck.Parsing.Common;
using Rubberduck.Parsing.VBA;
using Rubberduck.Properties;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
#if !DEBUG
    [Parsing.Common.Disabled]
#endif 
    public class RefactorExtractMethodCommandMenuItem : CommandMenuItemBase
    {
        public RefactorExtractMethodCommandMenuItem(CommandBase command) 
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
