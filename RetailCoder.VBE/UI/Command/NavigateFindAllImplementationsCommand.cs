using System;

namespace Rubberduck.UI.Command
{
    public class NavigateFindAllImplementationsCommand : ICommand
    {
        public void Execute()
        {
            throw new NotImplementedException();
        }
    }

    public class NavigateFindAllImplementationsCommandMenuItem : CommandMenuItemBase
    {
        public NavigateFindAllImplementationsCommandMenuItem(ICommand command) : base(command)
        {
        }

        public override string Key { get { return "ContextMenu_GoToImplementation"; } }
        public override int DisplayOrder { get { return (int)NavigationMenuItemDisplayOrder.FindImplementations; } }
    }
}