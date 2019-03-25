using System;
using System.Globalization;
using System.Windows.Data;
using Rubberduck.UI.ToDoItems;

namespace Rubberduck.UI.Converters
{
    public class ToDoItemGroupingToBooleanConverter : GroupingToBooleanConverter<ToDoItemGrouping> { }

    /// <summary>
    /// Provides a mutually exclusive binding between an ToDoItemGrouping and a boolean. 
    /// Note: This is a stateful converter, so each bound control requires its own converter instance.
    /// </summary>
    public class GroupingToBooleanConverter<T> : IValueConverter where T : IConvertible, IComparable
    {
        private T _state;

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (!(parameter is T governing) ||
                !(value is T bound))
            {
                return false;
            }

            _state = bound;
            return _state.Equals(governing);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (!(parameter is T governing) ||
                !(value is bool isSet))
            {
                return _state;
            }

            _state = isSet ? governing : _state;
            return _state;
        }
    }
}
