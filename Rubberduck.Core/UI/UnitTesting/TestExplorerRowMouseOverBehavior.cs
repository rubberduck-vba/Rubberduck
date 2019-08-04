using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Interactivity;
using GongSolutions.Wpf.DragDrop.Utilities;
using Rubberduck.UI.Controls;
using Rubberduck.UI.UnitTesting.ViewModels;

namespace Rubberduck.UI.UnitTesting
{
    public class TestExplorerRowMouseOverBehavior : Behavior<DataGrid>
    {
        public TestMethodViewModel MouseOverTest
        {
            get => GetValue(MouseOverTestProperty) as TestMethodViewModel;
            set => SetValue(MouseOverTestProperty, value);
        }

        public CollectionViewGroup MouseOverGroup
        {
            get => GetValue(MouseOverGroupProperty) as CollectionViewGroup;
            set => SetValue(MouseOverGroupProperty, value);
        }

        public static readonly DependencyProperty MouseOverTestProperty = DependencyProperty.Register(nameof(MouseOverTest),
            typeof(TestMethodViewModel), typeof(TestExplorerRowMouseOverBehavior));

        public static readonly DependencyProperty MouseOverGroupProperty = DependencyProperty.Register(nameof(MouseOverGroup),
            typeof(CollectionViewGroup), typeof(TestExplorerRowMouseOverBehavior));

        protected override void OnAttached()
        {
            base.OnAttached();

            var dataGrid = AssociatedObject;
            if (dataGrid == null)
            {
                return;
            }

            AddHandler(dataGrid);
        }

        protected override void OnDetaching()
        {
            var dataGrid = AssociatedObject;
            if (dataGrid != null)
            {
                RemoveHandler(dataGrid);
            }

            base.OnDetaching();
        }

        private void AddHandler(DataGrid dataGrid)
        {
            dataGrid.MouseMove += OnMouseMoved;
            dataGrid.ContextMenuOpening += OnContextMenuOpened;
            dataGrid.ContextMenuClosing += OnContextMenuClosed;
        }

        private void RemoveHandler(DataGrid dataGrid)
        {
            dataGrid.MouseMove -= OnMouseMoved;
            dataGrid.ContextMenuOpening -= OnContextMenuOpened;
            dataGrid.ContextMenuClosing -= OnContextMenuClosed;
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

        private void OnMouseMoved(object sender, MouseEventArgs e)
        {
            if (!(sender is GroupingGrid grid) || _contextMenuOpen)
            {
                return;
            }

            var isOverRow = false;
            foreach (var (row, model) in GetDataGridRows(grid))
            {
                if (!row.IsMouseOver)
                {
                    continue;
                }
                MouseOverTest = model;
                isOverRow = true;
                break;
            }

            if (!isOverRow)
            {
                MouseOverTest = null;
                // Over an expander?
                MouseOverGroup =
                    grid.GetVisualDescendents<Expander>().FirstOrDefault(group => group.IsMouseOver)?.DataContext as CollectionViewGroup;
            }
            else
            {
                MouseOverGroup = null;
            }
        }

        private static IEnumerable<(DataGridRow GridRow, TestMethodViewModel Model)> GetDataGridRows(ItemsControl grid)
        {
            var source = grid.ItemsSource;
            if (source is null)
            {
                yield break;
            }

            foreach (var item in source.OfType<TestMethodViewModel>())
            {
                if (grid.ItemContainerGenerator.ContainerFromItem(item) is DataGridRow row)
                {
                    yield return (row, item);
                }
            }
        }
    }
}
