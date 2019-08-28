using System;
using System.Globalization;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command;
using Rubberduck.Resources.CodeExplorer;
using Rubberduck.Formatters;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class CopyResultsCommand : CommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IClipboardWriter _clipboard;

        public CopyResultsCommand(RubberduckParserState state)
        {
            _state = state;
            _clipboard = new ClipboardWriter();

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            return _state.Status == ParserState.Ready;
        }

        protected override void OnExecute(object parameter)
        {
            ColumnInfo[] ColumnInfos = { new ColumnInfo("Project"), new ColumnInfo("Folder"), new ColumnInfo("Component"), new ColumnInfo("Declaration Type"), new ColumnInfo("Scope"),
                                       new ColumnInfo("Name"), new ColumnInfo("Return Type") };

            var declarationFormatters = _state.AllUserDeclarations.Select(declaration => new DeclarationFormatter(declaration));
            var title = string.Format(CodeExplorerUI.CodeExplorer_AppendHeader, DateTime.Now.ToString(CultureInfo.InvariantCulture));

            _clipboard.AppendInfo(ColumnInfos, declarationFormatters, title, ClipboardWriterAppendingInformationFormat.All);

            _clipboard.Flush();
        }
    }
}