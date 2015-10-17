using System.Collections.Generic;
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
        }

        public void Initialize()
        {
            foreach (var menu in _menus)
            {
                menu.Initialize();
            }
        }

        public void SetCommandButtonEnabledState(string key, bool isEnabled = true)
        {
            foreach (var menu in _menus)
            {
                menu.SetCommandButtonEnabledState(key, isEnabled);
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