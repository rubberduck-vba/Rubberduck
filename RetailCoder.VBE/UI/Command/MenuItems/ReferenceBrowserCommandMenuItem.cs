using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class ReferenceBrowserCommandMenuItem : CommandMenuItemBase
    {
        public ReferenceBrowserCommandMenuItem(CommandBase command) 
            : base(command)
        { }

        public override string Key { get { return "ReferenceBrowser_Menu"; } }
        public override int DisplayOrder { get { return (int) RubberduckMenuItemDisplayOrder.ReferenceBrowser; } }
    }
}
