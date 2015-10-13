using System.Collections.Generic;
using System.Diagnostics;
using Rubberduck.UI.Command;
using Rubberduck.UI.Command.MenuItems;
using Rubberduck.UI.Command.MenuItems.ParentMenus;

namespace Rubberduck
{
    public class AppMenu : IAppMenu
    {
        private readonly IEnumerable<IParentMenuItem> _menus;

        public AppMenu(IEnumerable<IParentMenuItem> menus)
        {
            _menus = menus;
            Debug.Print("in AppMenu ctor");
            foreach (var parentMenuItem in menus)
            {
                Debug.Print("'{0}' ({1})", parentMenuItem.Key, parentMenuItem.GetHashCode());
            }
        }

        public void Initialize()
        {
            foreach (var menu in _menus)
            {
                menu.Initialize();
            }
        }

        public void Localize()
        {
            foreach (var menu in _menus)
            {
                menu.Localize();
            }
        }
    }
}