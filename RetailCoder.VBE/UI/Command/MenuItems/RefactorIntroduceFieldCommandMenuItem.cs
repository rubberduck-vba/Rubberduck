using System.Drawing;
using Rubberduck.Parsing.VBA;
using Rubberduck.Properties;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class RefactorIntroduceFieldCommandMenuItem : CommandMenuItemBase
    {
        public RefactorIntroduceFieldCommandMenuItem (CommandBase command)
            : base(command)
        {
        }

        public override string Key => "RefactorMenu_IntroduceField";
        public override int DisplayOrder => (int)RefactoringsMenuItemDisplayOrder.IntroduceField;

        public override Image Image => Resources.AddVariable;
        public override Image Mask => Resources.AddVariableMask;


        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && Command.CanExecute(null);
        }
    }
}
