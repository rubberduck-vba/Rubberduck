using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interactivity;
using GongSolutions.Wpf.DragDrop.Utilities;

namespace Rubberduck.UI.Inspections
{
    public class InspectionContextMenuBehavior : Behavior<DataGrid>
    {
        protected override void OnAttached()
        {
            base.OnAttached();

            var element = AssociatedObject;
            if (element == null)
            {
                return;
            }

            AddHandler(element);
        }

        private void AddHandler(FrameworkElement element)
        {
            element.PreviewMouseRightButtonDown += OnClickStarting;
            element.MouseMove += OnMouseMoved;
            element.ContextMenuOpening += OnContextMenuOpened;
            element.ContextMenuClosing += OnContextMenuClosed;
        }

        private void RemoveHandler(FrameworkElement element)
        {
            element.PreviewMouseRightButtonDown -= OnClickStarting;
            element.MouseMove -= OnMouseMoved;
            element.ContextMenuOpening -= OnContextMenuOpened;
            element.ContextMenuClosing -= OnContextMenuClosed;
        }

        private void OnClickStarting(object sender, MouseButtonEventArgs e)
        {
            var element = AssociatedObject;
            var listItem = element?.GetVisualAncestor<ListViewItem>();

            if (listItem == null)
            {
                e.Handled = true;
                return;
            }

            listItem.IsSelected = true;
        }

        private void OnMouseMoved(object sender, MouseEventArgs e)
        {
            if (!(sender is DataGrid grid) || _contextMenuOpen || grid.ContextMenu is null)
            {
                return;
            }

            grid.ContextMenu.Visibility = GetDataGridRows(grid).Any(row => row.IsMouseOver) ? Visibility.Visible : Visibility.Collapsed;
        }

        private static IEnumerable<DataGridRow> GetDataGridRows(ItemsControl grid)
        {
            var source = grid.ItemsSource;
            if (source is null)
            {
                yield break;
            }

            foreach (var item in source)
            {
                if (grid.ItemContainerGenerator.ContainerFromItem(item) is DataGridRow row)
                {
                    yield return row;
                }
            }
        }
        
        private bool _contextMenuOpen;

        private void OnContextMenuOpened(object sender, EventArgs e)
        {
            _contextMenuOpen = true;
        }

        private void OnContextMenuClosed(object sender, EventArgs e)
        {
            _contextMenuOpen = false;
        }

        protected override void OnDetaching()
        {
            var listView = AssociatedObject;
            if (listView != null)
            {
                RemoveHandler(listView);
            }

            base.OnDetaching();
        }
    }
}
