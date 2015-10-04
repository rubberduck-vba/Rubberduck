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
    }

    public class ClipboardWriter : IClipboardWriter
    {
        public void Write(string text)
        {
            Clipboard.SetText(text);
        }
    }
}
