using System;
using System.Globalization;
using System.Windows.Data;
using Rubberduck.UnitTesting;

namespace Rubberduck.UI.UnitTesting
{
    public class TestResultToOutcomeTextConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is TestResult result)
            {
                return RubberduckUI.ResourceManager.GetString("TestOutcome_" + result.Outcome, CultureInfo.CurrentUICulture);
            }

            return null;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}