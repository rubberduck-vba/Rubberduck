using System;
using System.Collections.Generic;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;
using Rubberduck.AddRemoveReferences;
using ImageSourceConverter = Rubberduck.UI.Converters.ImageSourceConverter;

namespace Rubberduck.UI.AddRemoveReferences
{
    public class PriorityButtonImageSourceConverter : ImageSourceConverter, IMultiValueConverter
    {
        private enum IconKey
        {
            None,
            MoveUp,
            MoveUpDim,
            MoveDown,
            MoveDownDim
        }

        private readonly IDictionary<IconKey, ImageSource> _icons = new Dictionary<IconKey, ImageSource>
            {
                { IconKey.None, null },
                { IconKey.MoveUp , ToImageSource(Resources.RubberduckUI.arrow_090) },
                { IconKey.MoveUpDim, ToImageSource(Resources.RubberduckUI.arrow_090_dimmed) },
                { IconKey.MoveDown, ToImageSource(Resources.RubberduckUI.arrow_270) },
                { IconKey.MoveDownDim, ToImageSource(Resources.RubberduckUI.arrow_270_dimmed) }
            };

        public override object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return _icons[IconKey.None];
        }

        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (values is null ||
                values.Length != 2 ||
                !(values[1] is ReferenceModel reference) ||         //DataContext
                reference.IsBuiltIn ||
                !(parameter is string direction))
            {
                return _icons[IconKey.None];
            }

            var mouseOver = (bool)values[0];

            if (mouseOver)
            {
                return direction.Equals("Up") ? _icons[IconKey.MoveUp] : _icons[IconKey.MoveDown];
            }

            return direction.Equals("Up") ? _icons[IconKey.MoveUpDim] : _icons[IconKey.MoveDownDim];
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
