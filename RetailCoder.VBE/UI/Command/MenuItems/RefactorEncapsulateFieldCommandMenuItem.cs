using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class RefactorEncapsulateFieldCommandMenuItem : CommandMenuItemBase
    {
        public RefactorEncapsulateFieldCommandMenuItem(CommandBase command) 
            : base(command)
        {
        }

        public override string Key { get { return "RefactorMenu_EncapsulateField"; } }
        public override int DisplayOrder { get { return (int)RefactoringsMenuItemDisplayOrder.EncapsulateField; } }
        //public override Image Image { get { return Resources.AddProperty_5538_32; } }
        //public override Image Mask { get { return Resources.AddProperty_5538_321_Mask; } }

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && Command.CanExecute(null);
        }
    }
}
