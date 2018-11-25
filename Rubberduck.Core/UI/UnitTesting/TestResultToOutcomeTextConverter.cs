using System;
using System.Globalization;
using System.Windows.Data;
using Rubberduck.UnitTesting;
using Rubberduck.Resources.UnitTesting;

namespace Rubberduck.UI.UnitTesting
{
    public class TestResultToOutcomeTextConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is TestOutcome result)
            {
                return TestExplorer.ResourceManager.GetString($"TestOutcome_{result}", CultureInfo.CurrentUICulture);
            }

            return null;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}