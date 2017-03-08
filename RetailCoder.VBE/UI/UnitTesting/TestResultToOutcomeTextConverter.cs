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
            var result = value as TestResult;
            if (result != null)
            {
                return RubberduckUI.ResourceManager.GetString("TestOutcome_" + result.Outcome, CultureInfo.CurrentUICulture);
            }

            throw new ArgumentException("value is not a TestResult object.", "value");
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
