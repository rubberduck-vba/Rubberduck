using System;
using System.Globalization;
using System.Windows.Data;

namespace Rubberduck.UI.Inspections
{
    /// <summary>
    /// Provides a mutally exclusive binding between an InspectionResultGrouping and a boolean. 
    /// Note: This is a stateful converter, so each bound control requires its own converter instance.
    /// </summary>
    public class InspectionResultGroupingToBooleanConverter : IValueConverter
    {
        private InspectionResultGrouping _state;

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (!(parameter is InspectionResultGrouping governing) ||
                !(value is InspectionResultGrouping bound))
            {
                return false;
            }

            _state = bound;
            return _state == governing;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (!(parameter is InspectionResultGrouping governing) ||
                !(value is bool isSet))
            {
                return _state;
            }

            _state = isSet ? governing : _state;
            return _state;
        }
    }
}
