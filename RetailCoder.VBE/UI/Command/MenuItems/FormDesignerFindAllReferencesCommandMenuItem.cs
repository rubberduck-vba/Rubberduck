using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class FormDesignerFindAllReferencesCommandMenuItem : CommandMenuItemBase
    {
        public FormDesignerFindAllReferencesCommandMenuItem(CommandBase command)
            : base(command)
        {
        }

        public override bool BeginGroup { get { return true; } }
        public override string Key { get { return "ContextMenu_FindAllReferences"; } }
        public override int DisplayOrder { get { return (int)NavigationMenuItemDisplayOrder.FindAllReferences; } }

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {            
            return state != null && Command.CanExecute(null);
        }
    }
}
