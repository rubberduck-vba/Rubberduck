using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class RunSelectedTestMethodCommandMenuItem : CommandMenuItemBase
    {
        public RunSelectedTestMethodCommandMenuItem(RunSelectedTestMethodCommand command) : base(command) { }

        public override string Key => "ContextMenu_RunSelectedTest";
        public override int DisplayOrder => (int)CodePaneContextMenuItemDisplayOrder.RunSelectedTest;

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && Command.CanExecute(null);
        }

        public override bool HiddenWhenDisabled => true;

        public override bool BeginGroup => false;
    }
}
