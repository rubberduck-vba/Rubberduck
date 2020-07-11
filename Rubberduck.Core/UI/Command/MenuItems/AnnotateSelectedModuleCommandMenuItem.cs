using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems.ParentMenus;
using Rubberduck.UI.Command.Refactorings;

namespace Rubberduck.UI.Command.MenuItems
{
    public class AnnotateSelectedModuleCommandMenuItem : CommandMenuItemBase
    {
        public AnnotateSelectedModuleCommandMenuItem(AnnotateSelectedModuleCommand command)
            : base(command)
        { }

        public override string Key => "AnnotateMenu_SelectedModule";
        public override int DisplayOrder => (int)AnnotateParentMenuItemDisplayOrder.SelectedModule;

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && Command.CanExecute(null);
        }
    }
}