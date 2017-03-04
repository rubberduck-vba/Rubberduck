using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

    namespace Rubberduck.UI.Command.MenuItems
    {
        public class FindAllReferencesCommandMenuItem : CommandMenuItemBase
        {
            public FindAllReferencesCommandMenuItem(CommandBase command) 
                : base(command)
            {
            }

            public override string Key { get { return "ContextMenu_FindAllReferences"; } }
            public override int DisplayOrder { get { return (int)CodePaneContextMenuItemDisplayOrder.FindAllReferences; } }

            public override bool EvaluateCanExecute(RubberduckParserState state)
            {
                return state != null && Command.CanExecute(null);
            }
        }
    }
