using System.Windows.Input;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

    namespace Rubberduck.UI.Command.MenuItems
    {
        public class FindAllReferencesCommandMenuItem : CommandMenuItemBase
        {
            public FindAllReferencesCommandMenuItem(ICommand command) 
                : base(command)
            {
            }

            public override string Key { get { return "ContextMenu_FindAllReferences"; } }
            public override int DisplayOrder { get { return (int)NavigationMenuItemDisplayOrder.FindAllReferences; } }

            public override bool EvaluateCanExecute(RubberduckParserState state)
            {
                return state.Status == ParserState.Ready;
            }
        }
    }