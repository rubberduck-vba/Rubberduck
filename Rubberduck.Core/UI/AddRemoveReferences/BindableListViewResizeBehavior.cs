using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interactivity;
using GongSolutions.Wpf.DragDrop.Utilities;

namespace Rubberduck.UI.AddRemoveReferences
{
    public class BindableListViewResizeBehavior : Behavior<ListView>
    {
        protected override void OnAttached()
        {
            base.OnAttached();

            var listView = AssociatedObject;
            if (listView == null)
            {
                return;
            }

            AddHandler(listView);
        }

        private void AddHandler(FrameworkElement listView)
        {
            listView.SizeChanged += OnSizeChanged;
        }

        private void RemoveHandler(FrameworkElement listView)
        {
            listView.SizeChanged -= OnSizeChanged;
        }

        private void OnSizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (!e.WidthChanged)
            {
                return;
            }

            var listView = AssociatedObject;

            var textBlock = listView?.GetVisualDescendent<TextBlock>();
            // ListView uses a grid internally, so we need to overshoot and come back. WPWTF.
            var grid = textBlock?.GetVisualAncestor<Grid>();
            var target = grid?.ColumnDefinitions.FirstOrDefault(column => column.Width.IsStar);

            if (target == null)
            {
                return;
            }

            var scroll = listView.GetVisualDescendents<ScrollViewer>()
                .FirstOrDefault(element => element.ComputedVerticalScrollBarVisibility == Visibility.Visible);

            var scrollWidth = scroll != null && scroll.IsVisible ? scroll.ActualWidth - scroll.ViewportWidth : 0;
            var calculated = listView.ActualWidth - grid.ColumnDefinitions
                                 .Where(column => !ReferenceEquals(target, column))
                                 .Sum(column => column.Width.Value) - scrollWidth;

            target.Width = new GridLength(calculated);
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
