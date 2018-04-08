using System.Drawing;
using Rubberduck.Parsing.VBA;
using Rubberduck.Properties;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class RefactorExtractMethodCommandMenuItem : CommandMenuItemBase
    {
        public RefactorExtractMethodCommandMenuItem(CommandBase command) 
            : base(command)
        {
        }

        public override string Key => "RefactorMenu_ExtractMethod";
        public override int DisplayOrder => (int)RefactoringsMenuItemDisplayOrder.ExtractMethod;

        public override bool BeginGroup => true;

        public override Image Image => Resources.ExtractMethod;
        public override Image Mask => Resources.ExtractMethodMask;

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && Command.CanExecute(null);
        }
    }
}
