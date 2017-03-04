using System;
using System.Globalization;
using System.Linq;
using System.Windows.Data;

namespace Rubberduck.UI.SourceControl.Converters
{
    public class ChangeTypesToTextConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var values = value.ToString().Split(new[] {", "}, StringSplitOptions.RemoveEmptyEntries);
            var translatedValue = values.Select(s => RubberduckUI.ResourceManager.GetString("SourceControl_FileStatus_" + s));
            return string.Join(", ", translatedValue);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new InvalidOperationException();
        }
    }
}
