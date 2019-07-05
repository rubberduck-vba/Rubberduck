using System;
using System.Globalization;
using System.Windows.Data;

namespace Rubberduck.UI.Converters
{
    public class MillisecondToTimeMagnitudeConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (!(value is long milliseconds) || milliseconds == 0)
            {
                return string.Empty;
            }

            var time = TimeSpan.FromMilliseconds(milliseconds);

            if (time.TotalHours >= 1)
            {
                return $"{FormatAsRoundedFloat(time.TotalHours)} {Resources.UnitTesting.TestExplorer.TestOutcome_DurationHour}";
            }
            if (time.TotalMinutes >= 1)
            {
                return $"{FormatAsRoundedFloat(time.TotalMinutes)} {Resources.UnitTesting.TestExplorer.TestOutcome_DurationMinute}";
            }
            if (time.TotalSeconds >= 1)
            {
                return $"{FormatAsRoundedFloat(time.TotalSeconds)} {Resources.UnitTesting.TestExplorer.TestOutcome_DurationSecond}";
            }

            return $"{time.TotalMilliseconds:F0} {Resources.UnitTesting.TestExplorer.TestOutcome_DurationMillisecond}";
        }

        private string FormatAsRoundedFloat(double duration)
        {
            var rounded = Math.Round(duration, 2) * 100;

            if (rounded % 100 <= 0.001)
            {
                return $"{duration:F0}";
            }

            if (rounded % 10 <= 0.001)
            {
                return $"{duration:F1}";
            }

            return $"{duration:F2}";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
