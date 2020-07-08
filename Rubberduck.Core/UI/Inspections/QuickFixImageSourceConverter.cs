using System;
using System.Globalization;
using System.Windows.Media;
using Rubberduck.Resources.Inspections;
using ImageSourceConverter = Rubberduck.UI.Converters.ImageSourceConverter;

namespace Rubberduck.UI.Inspections
{
    public class QuickFixImageSourceConverter : ImageSourceConverter
    {
        private static readonly ImageSource IgnoreOnceIcon = ToImageSource(InspectionsUI.ignore_once);

        public override object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value != null 
                && value.GetType().Name.Equals("IgnoreOnceQuickFix"))
            {
                return IgnoreOnceIcon;
            }

            return null;
        }
    }
}
