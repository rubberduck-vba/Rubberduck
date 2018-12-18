using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interactivity;
using GongSolutions.Wpf.DragDrop.Utilities;

namespace Rubberduck.UI.Controls
{
    public class BindableListItemDrillThroughBehavior : Behavior<FrameworkElement>
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

        private void AddHandler(IInputElement element)
        {
            element.PreviewMouseLeftButtonDown += OnClickStarting;
        }

        private void RemoveHandler(IInputElement element)
        {
            element.PreviewMouseLeftButtonDown -= OnClickStarting;
        }

        private void OnClickStarting(object sender, MouseButtonEventArgs e)
        {
            var element = AssociatedObject;
            var listItem = element?.GetVisualAncestor<ListViewItem>();

            if (listItem == null)
            {
                return;
            }

            listItem.IsSelected = true;
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
