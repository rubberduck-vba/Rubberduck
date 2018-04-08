using System.Drawing;
using Rubberduck.Parsing.VBA;
using Rubberduck.Properties;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class RefactorRemoveParametersCommandMenuItem : CommandMenuItemBase
    {
        public RefactorRemoveParametersCommandMenuItem(CommandBase command) : base(command)
        {
        }

        public override string Key => "RefactorMenu_RemoveParameter";
        public override int DisplayOrder => (int)RefactoringsMenuItemDisplayOrder.RemoveParameters;
        public override bool BeginGroup => true;

        public override Image Image => Resources.RemoveParameters;
        public override Image Mask => Resources.RemoveParametersMask;

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && Command.CanExecute(null);
        }
    }
}
