using System.IO;
using System.Windows;
using System.Windows.Media.Imaging;
using System.Collections.Generic;
using System.Linq;
using System;

namespace Rubberduck.Common
{
    public interface IClipboardWriter
    {
        void Write(string text);
        void AppendImage(BitmapSource image);
        void AppendString(string formatName, string data);
        void AppendStream(string formatName, MemoryStream stream);
        void Flush();
        void AppendInfo<T>(ColumnInfo[] columnInfos,
            IEnumerable<T> exportableResults,
            string titleFormat,
            ClipboardWriterAppendingInformationFormat appendingInformationFormat) where T : IExportable;
    }

    public enum ClipboardWriterAppendingInformationFormat
    {
        XmlSpreadsheetFormat = 1 << 0,
        RtfFormat = 1 << 1,
        HtmlFormat = 1 << 2,
        CsvFormat = 1 << 3,
        UnicodeFormat = 1 << 4,
        All = XmlSpreadsheetFormat | RtfFormat | HtmlFormat | CsvFormat | UnicodeFormat
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

        public void AppendInfo<T>(ColumnInfo[] columnInfos,
            IEnumerable<T> results,
            string title,
            ClipboardWriterAppendingInformationFormat appendingInformationFormat) where T : IExportable
        {
            object[][] resultsAsArray = results.Select(result => result.ToArray()).ToArray();

            var includeXmlSpreadsheetFormat = (appendingInformationFormat & ClipboardWriterAppendingInformationFormat.XmlSpreadsheetFormat) == ClipboardWriterAppendingInformationFormat.XmlSpreadsheetFormat;
            if (includeXmlSpreadsheetFormat)
            {
                const string xmlSpreadsheetDataFormat = "XML Spreadsheet";
                using (var stream = ExportFormatter.XmlSpreadsheetNew(resultsAsArray, title, columnInfos))
                {
                    AppendStream(DataFormats.GetDataFormat(xmlSpreadsheetDataFormat).Name, stream);
                }
            }

            var includeRtfFormat = (appendingInformationFormat & ClipboardWriterAppendingInformationFormat.RtfFormat) == ClipboardWriterAppendingInformationFormat.RtfFormat;
            if (includeRtfFormat)
            {
                AppendString(DataFormats.Rtf, ExportFormatter.RTF(resultsAsArray, title));
            }

            var includeHtmlFormat = (appendingInformationFormat & ClipboardWriterAppendingInformationFormat.HtmlFormat) == ClipboardWriterAppendingInformationFormat.HtmlFormat;
            if (includeHtmlFormat)
            {
                AppendString(DataFormats.Html, ExportFormatter.HtmlClipboardFragment(resultsAsArray, title, columnInfos));
            }

            var includeCsvFormat = (appendingInformationFormat & ClipboardWriterAppendingInformationFormat.CsvFormat) == ClipboardWriterAppendingInformationFormat.CsvFormat;
            if (includeCsvFormat)
            {
                AppendString(DataFormats.CommaSeparatedValue, ExportFormatter.Csv(resultsAsArray, title, columnInfos));
            }

            var includeUnicodeFormat = (appendingInformationFormat & ClipboardWriterAppendingInformationFormat.UnicodeFormat) == ClipboardWriterAppendingInformationFormat.UnicodeFormat;
            if (includeUnicodeFormat && results is IEnumerable<IExportable> unicodeResults)
            {
                var unicodeTextFormat = title + Environment.NewLine + string.Join(string.Empty, unicodeResults.Select(result => result.ToClipboardString() + Environment.NewLine).ToArray());
                AppendString(DataFormats.UnicodeText, unicodeTextFormat);
            }
        }
    }
}
