using System.Windows;
using System.Windows.Controls;
using System.Windows.Interactivity;
using Rubberduck.Navigation.CodeExplorer;

namespace Rubberduck.UI.Controls
{
    /// <summary>
    /// http://stackoverflow.com/a/5118406/1188513
    /// </summary>
    public class BindableSelectedItemBehavior : Behavior<TreeView>
    {
        public object SelectedItem
        {
            get => (object) GetValue(SelectedItemProperty);
            set => SetValue(SelectedItemProperty, value);
        }

        public static readonly DependencyProperty SelectedItemProperty =
            DependencyProperty.Register("SelectedItem", typeof (object), typeof (BindableSelectedItemBehavior),
                new UIPropertyMetadata(null, OnSelectedItemChanged));

        private static void OnSelectedItemChanged(DependencyObject sender, DependencyPropertyChangedEventArgs e)
        {
            if (!(e.NewValue is CodeExplorerItemViewModel node) || 
                !(sender is BindableSelectedItemBehavior binding))
            {
                return;
            }

            var item = FindTreeViewItemFromData(binding.AssociatedObject, node);
            if (item == null)
            {
                return;
            }

            item.BringIntoView();
            item.Focus();
            item.IsSelected = true;
        }

        private static TreeViewItem FindTreeViewItemFromData(ItemsControl items, object node)
        {
            if (items.ItemContainerGenerator.ContainerFromItem(node) is TreeViewItem item)
            {
                return item;
            }

            foreach (var container in items.Items)
            {
                if (!(items.ItemContainerGenerator.ContainerFromItem(container) is TreeViewItem subItem))
                {
                    continue;
                }

                item = FindTreeViewItemFromData(subItem, node);

                if (item != null)
                {
                    return item;
                }
            }
            return null;
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
