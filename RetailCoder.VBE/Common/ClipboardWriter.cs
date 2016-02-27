using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace Rubberduck.Common
{
    public interface IClipboardWriter
    {
        void Write(string text);
        void AppendData(string format, object data);
        void Flush();
    }

    public class ClipboardWriter : IClipboardWriter
    {
        private DataObject _data;

        public void Write(string text)
        {
            this.AppendData(DataFormats.UnicodeText, text);
            this.Flush();
        }

        public void AppendData(string format, object data)
        {
            if (_data == null)
            {
                _data = new DataObject();
            }
            _data.SetData(format, data);
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
