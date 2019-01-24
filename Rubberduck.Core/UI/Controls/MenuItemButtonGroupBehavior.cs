using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Interactivity;

namespace Rubberduck.UI.Controls
{
    public class MenuItemButtonGroupBehavior : Behavior<MenuItem>
    {
        protected override void OnAttached()
        {
            base.OnAttached();

            GetCheckableSubMenuItems(AssociatedObject)
                .ToList()
                .ForEach(item => item.Click += OnClick);
        }

        protected override void OnDetaching()
        {
            base.OnDetaching();

            GetCheckableSubMenuItems(AssociatedObject)
                .ToList()
                .ForEach(item => item.Click -= OnClick);
        }

        private static IEnumerable<MenuItem> GetCheckableSubMenuItems(ItemsControl menuItem)
        {
            var itemCollection = menuItem.Items;
            return itemCollection.OfType<MenuItem>().Where(menuItemCandidate => menuItemCandidate.IsCheckable);
        }

        private void OnClick(object sender, RoutedEventArgs routedEventArgs)
        {
            var menuItem = (MenuItem)sender;

            if (!menuItem.IsChecked)
            {
                menuItem.IsChecked = true;
                return;
            }

            GetCheckableSubMenuItems(AssociatedObject)
                .Where(item => item != menuItem)
                .ToList()
                .ForEach(item => item.IsChecked = false);
        }
    }
}

