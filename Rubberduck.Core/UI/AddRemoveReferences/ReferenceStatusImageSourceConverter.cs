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
        private static readonly IDictionary<ReferenceStatus, ImageSource> Icons =
            new Dictionary<ReferenceStatus, ImageSource>
            {
                { ReferenceStatus.None, null },
                { ReferenceStatus.Pinned , ToImageSource(Resources.RubberduckUI.pinned) },
                { ReferenceStatus.Recent, ToImageSource(Resources.RubberduckUI.clock_select) },
                { ReferenceStatus.Recent | ReferenceStatus.Pinned, ToImageSource(Resources.RubberduckUI.clock_select_pinned) },
                { ReferenceStatus.BuiltIn, ToImageSource(Resources.RubberduckUI.padlock) },
                { ReferenceStatus.Broken, ToImageSource(Resources.RubberduckUI.exclamation) },
                { ReferenceStatus.Loaded, ToImageSource(Resources.RubberduckUI.tick_circle) },
                { ReferenceStatus.Added, ToImageSource(Resources.RubberduckUI.plus_circle) },
                { ReferenceStatus.BuiltIn | ReferenceStatus.Pinned, ToImageSource(Resources.RubberduckUI.lock_pinned) },
                { ReferenceStatus.Broken | ReferenceStatus.Pinned, ToImageSource(Resources.RubberduckUI.exclamation_pinned) },
                { ReferenceStatus.Loaded | ReferenceStatus.Pinned, ToImageSource(Resources.RubberduckUI.tick_circle_pinned) },
                { ReferenceStatus.Added | ReferenceStatus.Pinned, ToImageSource(Resources.RubberduckUI.plus_circle_pinned) }
            };

        public override object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return !(value is ReferenceStatus) ? null : Icons[(ReferenceStatus)value];
        }
    }
}