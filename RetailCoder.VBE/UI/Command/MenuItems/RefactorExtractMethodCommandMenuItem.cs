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

        public override string Key { get { return "RefactorMenu_ExtractMethod"; } }
        public override int DisplayOrder { get { return (int)RefactoringsMenuItemDisplayOrder.ExtractMethod; } }

        public override bool BeginGroup
        {
            get { return true; }
        }

        public override Image Image { get { return Resources.ExtractMethod; } }
        public override Image Mask { get { return Resources.ExtractMethodMask; } }

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && Command.CanExecute(null);
        }
    }
}
