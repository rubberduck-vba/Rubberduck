using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interactivity;

namespace Rubberduck.UI.Controls
{

    public class BindableSelectedItemBehavior : Behavior<TreeView>
    {
        public object SelectedItem
        {
            get => GetValue(SelectedItemProperty);
            set => SetValue(SelectedItemProperty, value);
        }

        public static readonly DependencyProperty SelectedItemProperty =
            DependencyProperty.Register("SelectedItem", typeof(object), typeof(BindableSelectedItemBehavior), new UIPropertyMetadata(null, OnSelectedItemChanged));

        private static void OnSelectedItemChanged(DependencyObject sender, DependencyPropertyChangedEventArgs e)
        {
            if (!(sender is BindableSelectedItemBehavior behavior) ||
                !(behavior.AssociatedObject is TreeView tree))
            {
                return;
            }

            if (e.NewValue == null)
            {
                foreach (var item in tree.Items.OfType<TreeViewItem>())
                {
                    item.SetValue(TreeViewItem.IsSelectedProperty, false);
                }
                return;
            }

            if (e.NewValue is TreeViewItem treeViewItem)
            {
                treeViewItem.SetValue(TreeViewItem.IsSelectedProperty, true);
            }
            else
            {
                var itemsHostProperty = tree.GetType().GetProperty("ItemsHost", BindingFlags.NonPublic | BindingFlags.Instance);
                if (itemsHostProperty == null)
                {
                    return;
                }

                if (!(itemsHostProperty.GetValue(tree, null) is Panel itemsHost))
                {
                    return;
                }
                foreach (var item in itemsHost.Children.OfType<TreeViewItem>())
                {
                    if (WalkTreeViewItem(item, e.NewValue))
                    {
                        break;
                    }
                }
            }
        }

        public static bool WalkTreeViewItem(TreeViewItem treeViewItem, object selectedValue)
        {
            if (treeViewItem.DataContext == selectedValue)
            {
                treeViewItem.SetValue(TreeViewItem.IsSelectedProperty, true);
                treeViewItem.Focus();
                return true;
            }

            var itemsHostProperty = treeViewItem.GetType().GetProperty("ItemsHost", BindingFlags.NonPublic | BindingFlags.Instance);

            if (itemsHostProperty == null ||
                !(itemsHostProperty.GetValue(treeViewItem, null) is Panel itemsHost))
            {
                return false;
            }

            foreach (var item in itemsHost.Children.OfType<TreeViewItem>())
            {
                if (WalkTreeViewItem(item, selectedValue))
                {
                    break;
                }
            }
            return false;
        }

        protected override void OnAttached()
        {
            base.OnAttached();
            AssociatedObject.SelectedItemChanged += OnTreeViewSelectedItemChanged;
        }

        protected override void OnDetaching()
        {
            base.OnDetaching();
            if (AssociatedObject != null)
            {
                AssociatedObject.SelectedItemChanged -= OnTreeViewSelectedItemChanged;
            }
        }

        private void OnTreeViewSelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            SelectedItem = e.NewValue;
        }
    }
}