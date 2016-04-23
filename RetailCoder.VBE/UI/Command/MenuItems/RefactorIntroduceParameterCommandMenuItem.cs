using System.Drawing;
using System.Windows.Input;
using Rubberduck.Parsing.VBA;
using Rubberduck.Properties;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class RefactorIntroduceParameterCommandMenuItem : CommandMenuItemBase
    {
        public RefactorIntroduceParameterCommandMenuItem (ICommand command)
            : base(command)
        {
        }

        public override string Key { get { return "RefactorMenu_IntroduceParameter"; } }
        public override int DisplayOrder { get { return (int)RefactoringsMenuItemDisplayOrder.IntroduceParameter; } }
        public override bool BeginGroup { get { return true; } }

        public override Image Image { get { return Resources.PromoteLocal_6784_32; } }
        public override Image Mask { get { return Resources.PromoteLocal_6784_32_Mask; } }

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return Command.CanExecute(null);
        }
    }
}