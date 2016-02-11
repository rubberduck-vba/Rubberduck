using System;
using System.Globalization;
using System.Windows.Data;

namespace Rubberduck.UI.Settings.Converters
{
    public class InspectionToToolTipConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var inspection = (InspectionSetting)value;
            var v = RubberduckUI.ResourceManager.GetString(inspection.Name + "Name") + Environment.NewLine +
                   RubberduckUI.ResourceManager.GetString(inspection.Name + "Meta");

            return RubberduckUI.ResourceManager.GetString(inspection.Name + "Name") + Environment.NewLine +
                   RubberduckUI.ResourceManager.GetString(inspection.Name + "Meta");
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value;
        }
    }
}