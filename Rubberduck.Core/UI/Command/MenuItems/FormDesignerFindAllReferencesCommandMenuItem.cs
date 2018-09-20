using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class FormDesignerFindAllReferencesCommandMenuItem : CommandMenuItemBase
    {
        public FormDesignerFindAllReferencesCommandMenuItem(FormDesignerFindAllReferencesCommand command)
            : base(command)
        {
        }

        public override bool BeginGroup => true;
        public override string Key => "ContextMenu_FindAllReferences";
        public override int DisplayOrder => (int)NavigationMenuItemDisplayOrder.FindAllReferences;

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {            
            return state != null && Command.CanExecute(null);
        }
    }
}
