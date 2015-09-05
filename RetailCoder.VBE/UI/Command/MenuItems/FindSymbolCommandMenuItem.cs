using System.Drawing;
using System.Windows.Input;
using Rubberduck.Properties;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck.UI.Command.MenuItems
{
    public class FindSymbolCommandMenuItem : CommandMenuItemBase
    {
        public FindSymbolCommandMenuItem(ICommand command) 
            : base(command)
        {
        }

        public override string Key { get {return "ContextMenu_FindSymbol"; } }
        public override int DisplayOrder { get { return (int)NavigationMenuItemDisplayOrder.FindSymbol; } }
        public override bool BeginGroup { get { return true; } }

        public override Image Image { get { return Resources.FindSymbol_6263_32; } }
        public override Image Mask { get { return Resources.FindSymbol_6263_32_Mask; } }
    }
}