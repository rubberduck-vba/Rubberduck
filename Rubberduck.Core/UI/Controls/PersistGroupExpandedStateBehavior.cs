using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interactivity;
using GongSolutions.Wpf.DragDrop.Utilities;

namespace Rubberduck.UI.Controls
{
    // Courtesy of @aKzenT - this is a heavily refactored implementation of https://stackoverflow.com/a/15924044/4088852
    public class PersistGroupExpandedStateBehavior : Behavior<Expander>
    {
        public static readonly DependencyProperty InitialExpandedStateProperty = DependencyProperty.Register(
            "InitialExpandedState",
            typeof(bool),
            typeof(PersistGroupExpandedStateBehavior),
            new PropertyMetadata(default(bool)));

        public static readonly DependencyProperty GroupNameProperty = DependencyProperty.Register(
            "GroupName",
            typeof(object),
            typeof(PersistGroupExpandedStateBehavior),
            new PropertyMetadata(default(object)));

        private static readonly DependencyProperty ExpandedStateStoreProperty =
            DependencyProperty.RegisterAttached(
                "ExpandedStateStore",
                typeof(IDictionary<object, bool>),
                typeof(PersistGroupExpandedStateBehavior),
                new PropertyMetadata(default(IDictionary<object, bool>)));

        public bool InitialExpandedState
        {
            get => (bool)GetValue(InitialExpandedStateProperty);
            set => SetValue(InitialExpandedStateProperty, value);
        }

        public object GroupName
        {
            get => GetValue(GroupNameProperty);
            set => SetValue(GroupNameProperty, value);
        }

        protected override void OnAttached()
        {
            base.OnAttached();

            var states = GetExpandedStateStore();
            var expanded = !states.ContainsKey(GroupName ?? string.Empty) ? (bool?)null : states[GroupName ?? string.Empty];

            if (expanded != null)
            {
                AssociatedObject.IsExpanded = InitialExpandedState;
            }

            AssociatedObject.Expanded += OnExpanded;
            AssociatedObject.Collapsed += OnCollapsed;
        }

        protected override void OnDetaching()
        {
            AssociatedObject.Expanded -= OnExpanded;
            AssociatedObject.Collapsed -= OnCollapsed;

            base.OnDetaching();
        }

        private void OnCollapsed(object sender, RoutedEventArgs e) => SetExpanded(false);

        private void OnExpanded(object sender, RoutedEventArgs e) => SetExpanded(true);

        private void SetExpanded(bool expanded)
        {
            var expandStateStore = GetExpandedStateStore();
            //TODO: Remove this once GetExpandedStateStore is reliable.
            if (expandStateStore != null)
            {
                expandStateStore[GroupName ?? string.Empty] = expanded;
            }
        }

        private IDictionary<object, bool> GetExpandedStateStore()
        {
            //FIXME: This is not reliable since the containing GroupItem does not necessarily have a VisualParent at the time of the request.
            if (!(AssociatedObject?.GetVisualAncestor<ItemsControl>() is ItemsControl items))
            {
                return null;
            }

            var states = (IDictionary<object, bool>)items.GetValue(ExpandedStateStoreProperty);

            if (states != null)
            {
                return states;
            }

            states = new Dictionary<object, bool>();
            items.SetValue(ExpandedStateStoreProperty, states);

            return states;
        }
    }
}
