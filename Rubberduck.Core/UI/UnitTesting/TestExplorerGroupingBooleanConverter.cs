using System;
using System.Globalization;
using System.Windows.Data;

namespace Rubberduck.UI.UnitTesting
{
    /// <summary>
    /// Binds an individual flag of <see cref="TestExplorerGrouping"/> to a boolean. Note: This is a stateful converter, so each bound control
    /// requires its own converter instance.
    /// </summary>
    internal class TestExplorerGroupingBooleanConverter : IValueConverter
    {
        private TestExplorerGrouping _state;

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (!(parameter is TestExplorerGrouping governing) ||
                !(value is TestExplorerGrouping bound))
            {
                return false;
            }

            _state = bound;
            return _state == governing;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (!(parameter is TestExplorerGrouping governing) ||
                !(value is bool isSet))
            {
                return _state;
            }

            _state = isSet ? governing : _state;
            return _state;
        }
    }
}
