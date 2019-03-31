using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Data;
using Rubberduck.Resources.Settings;
using Rubberduck.UnitTesting.Settings;


namespace Rubberduck.UI.Settings.Converters
{
    public class AssertModeToTextConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var modes = (IEnumerable<AssertMode>)value;
            return modes.Select(s => UnitTestingPage.ResourceManager.GetString("UnitTestSettings_" + s, CultureInfo.CurrentUICulture)).ToArray();
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value;
        }
    }
}
