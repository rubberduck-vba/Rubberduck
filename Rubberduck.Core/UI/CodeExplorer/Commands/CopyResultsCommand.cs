using System;
using System.Globalization;
using System.Linq;
using System.Windows;
using Rubberduck.Common;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.Command;

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

            //TODO: _state.AllUserDeclarations --> Results
            // this.ProjectName, this.CustomFolder, this.ComponentName, this.DeclarationType.ToString(), this.Scope 
            var aDeclarations = _state.AllUserDeclarations.Select(declaration => declaration.ToArray()).ToArray();

            const string resource = "Rubberduck User Declarations - {0}";
            _clipboard.AppendInfo(ColumnInfos, _state.AllUserDeclarations, resource, true, true, true, true);

            _clipboard.Flush();
        }
    }
}