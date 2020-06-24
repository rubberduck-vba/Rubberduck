using System;
using System.Globalization;
using System.Windows.Data;
using Rubberduck.Resources;

namespace Rubberduck.UI.Converters
{
    //Based on https://stackoverflow.com/a/29659265 by Yoh Deadfall
    public class EnumToLocalizedStringConverter : IValueConverter
    {
        public string ResourcePrefix { get; set; }

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null)
            {
                throw new ArgumentException("The value cannot be null.", "value");
            }

            //TODO: Make this independent of the resource used.
            return RubberduckUI.ResourceManager.GetString(ResourcePrefix + value.ToString(), culture);
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            var str = (string)value;

            foreach (var enumValue in Enum.GetValues(targetType))
            {
                //TODO: Make this independent of the resource used.
                if (str == RubberduckUI.ResourceManager.GetString(ResourcePrefix + enumValue.ToString(), culture))
                {
                    return enumValue;
                }
            }

            throw new ArgumentException("There is no enumeration member corresponding to the specified name.", "value");
        }
    }
}