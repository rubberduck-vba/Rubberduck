using System.Windows;
using System.Windows.Controls;

namespace Rubberduck.UI.Controls
{
    //https://stackoverflow.com/a/2212927/4088852
    public class EventFocusAttachment
    {
        public static Control GetElementToFocus(Button button)
        {
            return (Control)button.GetValue(ElementToFocusProperty);
        }

        public static void SetElementToFocus(Button button, Control value)
        {
            button.SetValue(ElementToFocusProperty, value);
        }

        public static readonly DependencyProperty ElementToFocusProperty =
            DependencyProperty.RegisterAttached("ElementToFocus", typeof(Control),
                typeof(EventFocusAttachment), new UIPropertyMetadata(null, ElementToFocusPropertyChanged));

        public static void ElementToFocusPropertyChanged(DependencyObject sender, DependencyPropertyChangedEventArgs e)
        {
            if (sender is Button button)
            {
                button.Click += (s, args) =>
                {
                    GetElementToFocus(button)?.Focus();
                };
            }
        }
    }
}
