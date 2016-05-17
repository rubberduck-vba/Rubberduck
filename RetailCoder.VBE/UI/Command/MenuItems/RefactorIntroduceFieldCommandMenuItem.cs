using System.Drawing;
using System.Windows.Input;
using Rubberduck.Parsing.VBA;
using Rubberduck.Properties;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class RefactorIntroduceFieldCommandMenuItem : CommandMenuItemBase
    {
        public RefactorIntroduceFieldCommandMenuItem (ICommand command)
            : base(command)
        {
        }

        public override string Key { get { return "RefactorMenu_IntroduceField"; } }
        public override int DisplayOrder { get { return (int)RefactoringsMenuItemDisplayOrder.IntroduceField; } }

        public override Image Image { get { return Resources.AddVariable_5541_32; } }
        public override Image Mask { get { return Resources.AddVariable_5541_32_Mask; } }


        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return Command.CanExecute(null);
        }
    }
}
