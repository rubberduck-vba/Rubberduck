using System.Windows;
using System.Windows.Controls.Primitives;

namespace Rubberduck.UI.Controls.Behavior
{
    using ButtonBehavior = System.Windows.Interactivity.Behavior<ButtonBase>;

    public class ResetValueBehavior : ButtonBehavior
    {
        private ButtonBase _associatedButton;

        protected override void OnAttached()
        {
            _associatedButton = AssociatedObject;

            _associatedButton.Click += AssociatedButtonClick;
        }

        protected override void OnDetaching()
        {
            _associatedButton.Click -= AssociatedButtonClick;
        }

        private void AssociatedButtonClick(object sender, RoutedEventArgs e)
        {
            Affects = Default;
        }

        public object Affects
        {
            get => GetValue(AffectsProperty);
            set
            {
                var oldValue = Affects;
                SetValue(AffectsProperty, value);
                OnPropertyChanged(new DependencyPropertyChangedEventArgs(AffectsProperty, oldValue, value));
            }
        }

        public object Default
        {
            get => GetValue(DefaultProperty);
            set
            {
                var oldValue = Default;
                SetValue(DefaultProperty, value);
                OnPropertyChanged(new DependencyPropertyChangedEventArgs(DefaultProperty, oldValue, value));
            }
        }

        // Using a DependencyProperty as the backing store for Affects. 
        // This enables animation, styling, binding, etc...
        public static readonly DependencyProperty AffectsProperty =
            DependencyProperty.Register("Affects", typeof(object), typeof(ResetValueBehavior), new UIPropertyMetadata());
        public static readonly DependencyProperty DefaultProperty =
            DependencyProperty.Register("Default", typeof(object), typeof(ResetValueBehavior), new UIPropertyMetadata(defaultValue: null));
    }
}
