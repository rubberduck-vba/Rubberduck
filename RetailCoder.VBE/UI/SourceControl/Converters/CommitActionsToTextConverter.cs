using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Data;

namespace Rubberduck.UI.SourceControl.Converters
{
    public class CommitActionsToTextConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var modes = (IEnumerable<CommitAction>)value;
            return modes.Select(s => RubberduckUI.ResourceManager.GetString("SourceControl_" + s, CultureInfo.CurrentUICulture)).ToArray();
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return value;
        }
    }
}
