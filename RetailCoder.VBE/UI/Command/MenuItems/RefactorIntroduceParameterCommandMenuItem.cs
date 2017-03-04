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

        public override string Key { get { return "RefactorMenu_IntroduceParameter"; } }
        public override int DisplayOrder { get { return (int)RefactoringsMenuItemDisplayOrder.IntroduceParameter; } }
        public override bool BeginGroup { get { return true; } }

        public override Image Image { get { return Resources.PromoteLocal; } }
        public override Image Mask { get { return Resources.PromoteLocalMask; } }

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && Command.CanExecute(null);
        }
    }
}
