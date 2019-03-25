using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class RunSelectedTestModuleCommandMenuItem : CommandMenuItemBase
    {
        public RunSelectedTestModuleCommandMenuItem(RunSelectedTestModuleCommand command) : base(command) { }

        public override string Key => "ContextMenu_RunSelectedTestModule";
        public override int DisplayOrder => (int)CodePaneContextMenuItemDisplayOrder.RunSelectedTestModule;

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return state != null && Command.CanExecute(null);
        }

        public override bool HiddenWhenDisabled => true;

        public override bool BeginGroup => true;
    }
}
