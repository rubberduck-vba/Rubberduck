using System.IO;
using System.Windows;
using System.Windows.Media.Imaging;

namespace Rubberduck.Common
{
    public interface IClipboardWriter
    {
        void Write(string text);
        void AppendImage(BitmapSource image);
        void AppendString(string formatName, string data);
        void AppendStream(string formatName, MemoryStream stream);
        void Flush();
    }

    public class ClipboardWriter : IClipboardWriter
    {
        private DataObject _data;

        public void Write(string text)
        {
            this.AppendString(DataFormats.UnicodeText, text);
            this.Flush();
        }

        public void AppendImage(BitmapSource image)
        {
            if (_data == null)
            {
                _data = new DataObject();
            }
            _data.SetImage(image);
        }


        public void AppendString(string formatName, string data)
        {
            if (_data == null)
            {
                _data = new DataObject();
            }
            _data.SetData(formatName, data);
        }

        public void AppendStream(string formatName, MemoryStream stream)
        {
            if (_data == null)
            {
                _data = new DataObject();
            }
            _data.SetData(formatName, stream);
        }
        
        public void Flush()
        {
            if (_data != null)
            {
                Clipboard.SetDataObject(_data, true);
                _data = null;
            }
        }
    }
}
