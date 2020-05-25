using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems.ParentMenus;
using Rubberduck.UI.Command.Refactorings;

namespace Rubberduck.UI.Command.MenuItems
{
    public class AnnotateSelectedDeclarationCommandMenuItem : CommandMenuItemBase
    {
        public AnnotateSelectedDeclarationCommandMenuItem(AnnotateSelectedDeclarationCommand command)
            : base(command)
        { }

        public override string Key => "AnnotateMenu_SelectedDeclaration";
        public override int DisplayOrder => (int)AnnotateParentMenuItemDisplayOrder.SelectedDeclaration;

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && Command.CanExecute(null);
        }
    }
}