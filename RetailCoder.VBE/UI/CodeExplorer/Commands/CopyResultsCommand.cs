using System;
using System.Globalization;
using System.Linq;
using System.Windows;
using NLog;
using Rubberduck.Common;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    [CodeExplorerCommand]
    public class CopyResultsCommand : CommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IClipboardWriter _clipboard;

        public CopyResultsCommand(RubberduckParserState state) : base(LogManager.GetCurrentClassLogger())
        {
            _state = state;
            _clipboard = new ClipboardWriter();
        }

        protected override void ExecuteImpl(object parameter)
        {
                const string XML_SPREADSHEET_DATA_FORMAT = "XML Spreadsheet";

                ColumnInfo[] ColumnInfos = { new ColumnInfo("Project"), new ColumnInfo("Folder"), new ColumnInfo("Component"), new ColumnInfo("Declaration Type"), new ColumnInfo("Scope"),
                                           new ColumnInfo("Name"), new ColumnInfo("Return Type") };

                // this.ProjectName, this.CustomFolder, this.ComponentName, this.DeclarationType.ToString(), this.Scope 
                var aDeclarations = _state.AllUserDeclarations.Select(declaration => declaration.ToArray()).ToArray();

                const string resource = "Rubberduck User Declarations - {0}";
                var title = string.Format(resource, DateTime.Now.ToString(CultureInfo.InvariantCulture));
            
                var csvResults = ExportFormatter.Csv(aDeclarations, title, ColumnInfos);
                var htmlResults = ExportFormatter.HtmlClipboardFragment(aDeclarations, title, ColumnInfos);
                var rtfResults = ExportFormatter.RTF(aDeclarations, title);

                var strm1 = ExportFormatter.XmlSpreadsheetNew(aDeclarations, title, ColumnInfos);
                //Add the formats from richest formatting to least formatting
                _clipboard.AppendStream(DataFormats.GetDataFormat(XML_SPREADSHEET_DATA_FORMAT).Name, strm1);
                _clipboard.AppendString(DataFormats.Rtf, rtfResults);
                _clipboard.AppendString(DataFormats.Html, htmlResults);
                _clipboard.AppendString(DataFormats.CommaSeparatedValue, csvResults);

                _clipboard.Flush();
        }
    }
}