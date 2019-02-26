using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.ComCommands;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public abstract class AddRemoveReferencesCommandMenuItemBase : CommandMenuItemBase
    {
        protected AddRemoveReferencesCommandMenuItemBase(AddRemoveReferencesCommand command) : base(command) { }

        public override string Key => "AddRemoveReferences";
        public override bool BeginGroup => true;

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && Command.CanExecute(null);
        }
    }

    public class ToolMenuAddRemoveReferencesCommandMenuItem : AddRemoveReferencesCommandMenuItemBase
    {
        public override int DisplayOrder => (int)ToolsMenuItemDisplayOrder.AddRemoveReferences;

        public ToolMenuAddRemoveReferencesCommandMenuItem(AddRemoveReferencesCommand command) : base(command) { }
    }

    public class ProjectExplorerAddRemoveReferencesCommandMenuItem : AddRemoveReferencesCommandMenuItemBase
    {
        public override int DisplayOrder => (int)RefactoringsMenuItemDisplayOrder.AddRemoveReferences;

        public ProjectExplorerAddRemoveReferencesCommandMenuItem(AddRemoveReferencesCommand command) : base(command) { }
    }
}
