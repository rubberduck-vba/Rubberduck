using System.Drawing;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems.ParentMenus;
using Rubberduck.UI.Command.Refactorings;

namespace Rubberduck.UI.Command.MenuItems
{
    public class RefactorIntroduceFieldCommandMenuItem : CommandMenuItemBase
    {
        public RefactorIntroduceFieldCommandMenuItem (RefactorIntroduceFieldCommand command)
            : base(command)
        {
        }

        public override string Key => "RefactorMenu_IntroduceField";
        public override int DisplayOrder => (int)RefactoringsMenuItemDisplayOrder.IntroduceField;

        public override Image Image => Resources.CommandBarIcons.AddVariable;
        public override Image Mask => Resources.CommandBarIcons.AddVariableMask;


        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && Command.CanExecute(null);
        }
    }
}
