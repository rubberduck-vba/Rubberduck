using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Windows.Data;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace Rubberduck.UI.Converters
{
    public abstract class ImageSourceConverter : IValueConverter
    {
        protected static ImageSource ToImageSource(Image source)
        {
            using (var ms = new MemoryStream())
            {
                ((Bitmap) source).Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                var image = new BitmapImage();
                image.BeginInit();
                image.CacheOption = BitmapCacheOption.OnLoad;
                ms.Seek(0, SeekOrigin.Begin);
                image.StreamSource = ms;
                image.EndInit();
                image.Freeze();

                return image;
            }
        }

        public abstract object Convert(object value, Type targetType, object parameter, CultureInfo culture);

        public virtual object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}