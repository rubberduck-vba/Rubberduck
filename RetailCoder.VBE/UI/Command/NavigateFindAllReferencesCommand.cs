using System;

namespace Rubberduck.UI.Command
{
    public class NavigateFindAllReferencesCommand : ICommand
    {
        public void Execute()
        {
            throw new NotImplementedException();
        }
    }

    public class NavigateFindAllReferencesCommandMenuItem : CommandMenuItemBase
    {
        public NavigateFindAllReferencesCommandMenuItem(ICommand command) 
            : base(command)
        {
        }

        public override string Key { get { return "ContextMenu_FindAllReferences"; } }
        public override int DisplayOrder { get { return (int)NavigationMenuItemDisplayOrder.FindAllReferences; } }
    }
}