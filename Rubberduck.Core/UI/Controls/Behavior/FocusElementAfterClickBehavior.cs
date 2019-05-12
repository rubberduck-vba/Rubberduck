using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Interactivity;

namespace Rubberduck.UI.Controls
{
    //https://stackoverflow.com/a/5161792/4088852
    public class FocusElementAfterClickBehavior : Behavior<ButtonBase>
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
            Keyboard.Focus(FocusElement);
        }

        public Control FocusElement
        {
            get => (Control)GetValue(FocusElementProperty);
            set => SetValue(FocusElementProperty, value);
        }

        // Using a DependencyProperty as the backing store for FocusElement.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty FocusElementProperty =
            DependencyProperty.Register("FocusElement", typeof(Control), typeof(FocusElementAfterClickBehavior), new UIPropertyMetadata());
    }
}
