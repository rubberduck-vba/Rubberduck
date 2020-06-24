using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems.ParentMenus;
using Rubberduck.UI.Command.Refactorings;

namespace Rubberduck.UI.Command.MenuItems
{
    public class AnnotateSelectedMemberCommandMenuItem : CommandMenuItemBase
    {
        public AnnotateSelectedMemberCommandMenuItem(AnnotateSelectedMemberCommand command)
            : base(command)
        { }

        public override string Key => "AnnotateMenu_SelectedMember";
        public override int DisplayOrder => (int)AnnotateParentMenuItemDisplayOrder.SelectedMember;

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && Command.CanExecute(null);
        }
    }
}