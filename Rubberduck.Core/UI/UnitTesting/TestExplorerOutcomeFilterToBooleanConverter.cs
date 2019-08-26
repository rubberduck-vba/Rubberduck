using System;
using System.Globalization;
using System.Windows.Data;

namespace Rubberduck.UI.UnitTesting
{
    /// <summary>
    /// Binds an individual flag of <see cref="TestExplorerOutcomeFilter"/> to a boolean. Note: This is a stateful converter, so each bound control
    /// requires its own converter instance.
    /// </summary>
    class TestExplorerOutcomeFilterToBooleanConverter : IValueConverter
    {
        private TestExplorerOutcomeFilter _cachedOutcomeFilter;

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (!(parameter is TestExplorerOutcomeFilter outcomeParameter)
                || !(value is TestExplorerOutcomeFilter outcomeCurrentlyFiltering))
            {
                return false;
            }

            _cachedOutcomeFilter = outcomeCurrentlyFiltering;
            return _cachedOutcomeFilter.HasFlag(outcomeParameter);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (!(parameter is TestExplorerOutcomeFilter outcomeParameter)
                || !(value is bool isApplied))
            {
                return _cachedOutcomeFilter;
            }

            _cachedOutcomeFilter = isApplied
                ? _cachedOutcomeFilter | outcomeParameter
                : _cachedOutcomeFilter ^ outcomeParameter;
            return _cachedOutcomeFilter;
        }
    }
}
