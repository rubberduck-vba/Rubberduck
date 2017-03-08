using System;
using System.Globalization;
using System.Windows.Data;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.UI.Controls
{
    public class DeclarationTypeToStringConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is DeclarationType)
            {
                var type = (DeclarationType)value;
                var text = RubberduckUI.ResourceManager.GetString("DeclarationType_" + type, CultureInfo.CurrentUICulture) ?? string.Empty;
                return CultureInfo.CurrentCulture.TextInfo.ToTitleCase(text);
            }

            return null;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
