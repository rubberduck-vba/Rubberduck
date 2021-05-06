using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.ComCommands;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class PeekDefinitionCommandMenuItem : CommandMenuItemBase
    {
        public PeekDefinitionCommandMenuItem(PeekDefinitionCommand command)
            : base(command)
        {}

        public override string Key => "ContextMenu_PeekDefinition";
        public override bool BeginGroup => true;

        public override int DisplayOrder => (int) CodePaneContextMenuItemDisplayOrder.PeekDefinition;

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state?.Status == ParserState.Ready;
        }
    }

    public class ProjectExplorerPeekDefinitionCommandMenuItem : PeekDefinitionCommandMenuItem
    {
        public ProjectExplorerPeekDefinitionCommandMenuItem(ProjectExplorerPeekDefinitionCommand command)
            : base(command)
        {}

        public override int DisplayOrder => (int) ProjectExplorerContextMenuItemDisplayOrder.PeekDefinition;
    }
}