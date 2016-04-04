using System.Drawing;
using System.Windows.Input;
using Rubberduck.Parsing.VBA;
using Rubberduck.Properties;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class RefactorEncapsulateFieldCommandMenuItem : CommandMenuItemBase
    {
        public RefactorEncapsulateFieldCommandMenuItem(ICommand command) 
            : base(command)
        {
        }

        public override string Key { get { return "RefactorMenu_EncapsulateField"; } }
        public override int DisplayOrder { get { return (int)RefactoringsMenuItemDisplayOrder.EncapsulateField; } }
        public override Image Image { get { return Resources.AddProperty_5538_32; } }

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state.Status == ParserState.Ready;
        }
    }
}