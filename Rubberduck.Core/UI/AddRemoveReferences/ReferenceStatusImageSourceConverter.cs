using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Media;
using Rubberduck.AddRemoveReferences;
using ImageSourceConverter = Rubberduck.UI.Converters.ImageSourceConverter;

namespace Rubberduck.UI.AddRemoveReferences
{
    public class ReferenceStatusImageSourceConverter : ImageSourceConverter
    {
        private readonly IDictionary<ReferenceStatus, ImageSource> _icons =
            new Dictionary<ReferenceStatus, ImageSource>
            {
                { ReferenceStatus.BuiltIn, ToImageSource(Resources.RubberduckUI.tick_circle) },
                { ReferenceStatus.Broken, ToImageSource(Resources.RubberduckUI.exclamation) },
                { ReferenceStatus.Loaded, ToImageSource(Resources.RubberduckUI.tick_circle) },
                { ReferenceStatus.Removed, ToImageSource(Resources.RubberduckUI.cross_circle) }
            };

        public override object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return _icons.First(); // todo fix this: wrecks the xaml designer otherwise

            if (value == null) { return null; }
            if (value.GetType() != typeof(ReferenceStatus))
            {
                throw new ArgumentException("value must be a ReferenceStatus");
            }

            var status = (ReferenceStatus)value;
            return _icons[status];
        }
    }
}