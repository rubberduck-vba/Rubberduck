using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.ComCommands;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class FindAllReferencesCommandMenuItem : CommandMenuItemBase
    {
        public FindAllReferencesCommandMenuItem(FindAllReferencesCommand command)
            : base(command)
        {
        }

        public override string Key => "ContextMenu_FindAllReferences";
        public override int DisplayOrder => (int) CodePaneContextMenuItemDisplayOrder.FindAllReferences;

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && Command.CanExecute(null);
        }
    }
}
