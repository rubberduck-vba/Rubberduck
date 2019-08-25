using System.IO;
using System.Windows;
using System.Windows.Media.Imaging;
using System.Collections.Generic;
using System.Linq;
using System;
using System.Globalization;
using System.Windows.Data;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI.UnitTesting.ViewModels;
using Rubberduck.Resources;

namespace Rubberduck.Common
{
    public interface IClipboardWriter
    {
        void Write(string text);
        void AppendImage(BitmapSource image);
        void AppendString(string formatName, string data);
        void AppendStream(string formatName, MemoryStream stream);
        void Flush();
        void AppendInfo(ColumnInfo[] columnInfos,
            object results, 
            string titleFormat,
            bool includeXmlSpreadsheetFormat = false, 
            bool includeRtfFormat = false, 
            bool includeHtmlFormat = false, 
            bool includeCsvFormat = false, 
            bool includeUnicodeFormat = false);
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
        public void AppendInfo(ColumnInfo[] columnInfos, 
            object results, 
            string titleFormat,
            bool includeXmlSpreadsheetFormat = false,
            bool includeRtfFormat = false,
            bool includeHtmlFormat = false, 
            bool includeCsvFormat = false, 
            bool includeUnicodeFormat = false)
        {
            //var resultsAsArray = results.Select(result => result.ToArray()).ToArray();
            object[][] resultsAsArray;
            switch (results)
            {
                case IEnumerable<Declaration> declarations:
                    resultsAsArray = declarations.Select(declaration => declaration.ToArray()).ToArray();
                    break;
                case System.Collections.ObjectModel.ObservableCollection<IInspectionResult> inspectionResults:
                    resultsAsArray = inspectionResults.OfType<IExportable>().Select(result => result.ToArray()).ToArray();
                    break;
                case System.Collections.ObjectModel.ObservableCollection<TestMethodViewModel> testMethodViewModels:
                    resultsAsArray = testMethodViewModels.Select(test => test.ToArray()).ToArray();
                    break;
                default:
                    resultsAsArray = null;
                    break;
            }
            
            var title = string.Format(RubberduckUI.TestExplorer_AppendHeader, DateTime.Now.ToString(CultureInfo.InvariantCulture));

            if (includeXmlSpreadsheetFormat)
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

            if (includeUnicodeFormat && results is ListCollectionView unicodeResults)
            {
                var unicodeTextFormat = title + Environment.NewLine + string.Join(string.Empty, unicodeResults.OfType<IExportable>().Select(result => result.ToClipboardString() + Environment.NewLine).ToArray());
                AppendString(DataFormats.UnicodeText, unicodeTextFormat);
            }
        }
    }
}
