using System.Windows;
using System.Windows.Controls;
using System.Windows.Interactivity;
using GongSolutions.Wpf.DragDrop.Utilities;

namespace Rubberduck.UI.Controls
{
    public class GroupItemExpandedBehavior : Behavior<DataGrid>
    {
        public object ExpandedState
        {
            get => GetValue(ExpandedStateProperty);
            set => SetValue(ExpandedStateProperty, value);
        }

        public static readonly DependencyProperty ExpandedStateProperty =
            DependencyProperty.Register("ExpandedState", typeof(object), typeof(GroupItemExpandedBehavior), new UIPropertyMetadata(null, OnExpandedStateChanged));

        private static void OnExpandedStateChanged(DependencyObject sender, DependencyPropertyChangedEventArgs e)
        {
            if (!(sender is GroupItemExpandedBehavior data) ||
                !(data.AssociatedObject is DataGrid grid) ||
                !grid.IsGrouping ||
                !(e.NewValue is bool))
            {
                return;
            }

            foreach (var expander in grid.GetVisualDescendents<Expander>())
            {
                expander.IsExpanded = (bool)e.NewValue;
            }
        }
    }
}
