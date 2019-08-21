using System.IO;
using System.Windows;
using System.Windows.Media.Imaging;
using System.Collections.Generic;
using System.Linq;
using System;
using System.Globalization;

namespace Rubberduck.Common
{
    public interface IClipboardWriter
    {
        void Write(string text);
        void AppendImage(BitmapSource image);
        void AppendString(string formatName, string data);
        void AppendStream(string formatName, MemoryStream stream);
        void Flush();
        //void AppendInfo(ClipboardFormat xmlSpreadsheetFormat, ClipboardFormat rtfFormat, ClipboardFormat htmlFormat, ClipboardFormat csvFormat, ClipboardFormat unicodeTextFormat);
        void AppendInfo(ColumnInfo[] columnInfos, IEnumerable<object> results, string titleFormat,
            bool includeXmlSpreadsheetformat = false, bool includeRtfFormat = false, bool includeHtmlFormat = false, bool includeCsvFormat = false, bool includeUnicodeFormat = false)
    }

    public struct ClipboardFormat
    {
        public string FormatName;
        public object Data;

        public ClipboardFormat(string formatName, object data)
        {
            FormatName = formatName;
            Data = data;
        }
    }

    public class ClipboardWriter : IClipboardWriter
    {
        private DataObject _data;

        public void Write(string text)
        {
            AppendString(DataFormats.UnicodeText, text);
            Flush();
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

        //public void AppendInfo(ClipboardFormat xmlSpreadsheetFormat, ClipboardFormat rtfFormat, ClipboardFormat htmlFormat, ClipboardFormat csvFormat, ClipboardFormat unicodeTextFormat)
        //TODO: bitFlag
        public void AppendInfo(ColumnInfo[] columnInfos, IEnumerable<object> results, 
            string titleFormat,
            bool includeXmlSpreadsheetformat = false,
            bool includeRtfFormat = false,
            bool includeHtmlFormat = false, 
            bool includeCsvFormat = false, 
            bool includeUnicodeFormat = false)
        {
            var resultsAsArray = results.Select(result => result.ToArray()).ToArray();
            var title = string.Format(titleFormat, DateTime.Now.ToString(CultureInfo.InvariantCulture));

            if (includeXmlSpreadsheetformat)
            {
                const string xmlSpreadsheetDataFormat = "XML Spreadsheet";
                using (var stream = ExportFormatter.XmlSpreadsheetNew(resultsAsArray, title, columnInfos))
                {
                    AppendStream(DataFormats.GetDataFormat(xmlSpreadsheetDataFormat).Name, stream);
                }
            }

            if (includeRtfFormat)
            {
                AppendString(DataFormats.Rtf, ExportFormatter.RTF(resultsAsArray, title));
            }

            if (includeHtmlFormat)
            {
                AppendString(DataFormats.Html, ExportFormatter.HtmlClipboardFragment(resultsAsArray, title, columnInfos));
            }

            if (includeCsvFormat)
            {
                AppendString(DataFormats.CommaSeparatedValue, ExportFormatter.Csv(resultsAsArray, title, columnInfos));
            }

            if (includeUnicodeFormat)
            {
                var unicodeResults = title + Environment.NewLine + string.Join(string.Empty, results.OfType<IExportable>().Select(result => result.ToClipboardString() + Environment.NewLine).ToArray());
                var unicodeTextFormat = new ClipboardFormat(DataFormats.UnicodeText, unicodeResults);
            }
        }
    }
}
