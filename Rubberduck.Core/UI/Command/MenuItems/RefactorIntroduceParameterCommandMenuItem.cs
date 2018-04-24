using System.Drawing;
using Rubberduck.Parsing.VBA;
using Rubberduck.Properties;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class RefactorIntroduceParameterCommandMenuItem : CommandMenuItemBase
    {
        public RefactorIntroduceParameterCommandMenuItem (CommandBase command)
            : base(command)
        {
        }

        public override string Key => "RefactorMenu_IntroduceParameter";
        public override int DisplayOrder => (int)RefactoringsMenuItemDisplayOrder.IntroduceParameter;
        public override bool BeginGroup => true;

        public override Image Image => Resources.PromoteLocal;
        public override Image Mask => Resources.PromoteLocalMask;

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && Command.CanExecute(null);
        }
    }
}
