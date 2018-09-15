using System.Drawing;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems.ParentMenus;
using Rubberduck.UI.Command.Refactorings;

namespace Rubberduck.UI.Command.MenuItems
{
    public class RefactorIntroduceParameterCommandMenuItem : CommandMenuItemBase
    {
        public RefactorIntroduceParameterCommandMenuItem (RefactorIntroduceParameterCommand command)
            : base(command)
        {
        }

        public override string Key => "RefactorMenu_IntroduceParameter";
        public override int DisplayOrder => (int)RefactoringsMenuItemDisplayOrder.IntroduceParameter;
        public override bool BeginGroup => true;

        public override Image Image => Resources.CommandBarIcons.PromoteLocal;
        public override Image Mask => Resources.CommandBarIcons.PromoteLocalMask;

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && Command.CanExecute(null);
        }
    }
}
