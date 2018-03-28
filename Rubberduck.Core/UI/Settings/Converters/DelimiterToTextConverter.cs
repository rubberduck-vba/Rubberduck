﻿using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Data;

namespace Rubberduck.UI.Settings.Converters
{
    public class DelimiterToTextConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var modes = (IEnumerable<DelimiterOptions>)value;
            return modes.Select(s => RubberduckUI.ResourceManager.GetString("GeneralSettings_" + s + "Delimiter", CultureInfo.CurrentUICulture)).ToArray();
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value;
        }
    }
}
