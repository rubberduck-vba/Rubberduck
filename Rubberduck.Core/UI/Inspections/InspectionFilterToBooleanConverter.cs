using System;
using System.Globalization;
using System.Windows.Data;

namespace Rubberduck.UI.Inspections
{
    internal class InspectionFilterToBooleanConverter : IValueConverter
    {
        private InspectionResultsFilter _state;

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (!(parameter is InspectionResultsFilter flag) ||
                !(value is InspectionResultsFilter bound))
            {
                return _state;
            }

            _state = bound;
            return _state.HasFlag(flag);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (!(parameter is InspectionResultsFilter flag) ||
                !(value is bool isSet))
            {
                return _state;
            }

            _state ^= flag;
            return _state;
        }       
    }
}
