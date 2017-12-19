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

        public override string Key => "RefactorMenu_EncapsulateField";
        public override int DisplayOrder => (int)RefactoringsMenuItemDisplayOrder.EncapsulateField;

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && Command.CanExecute(null);
        }
    }
}
