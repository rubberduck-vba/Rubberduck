using System;
using Rubberduck.Parsing.VBA.Parsing.ParsingExceptions;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Parsing.VBA
{
    public class ParseErrorEventArgs : EventArgs
    {
        private readonly QualifiedModuleName _moduleName;

        public ParseErrorEventArgs(SyntaxErrorException exception, QualifiedModuleName moduleName)
        {
            Exception = exception;
            _moduleName = moduleName;
        }

        public SyntaxErrorException Exception { get; }

        public string ComponentName => _moduleName.ComponentName;
        public string ProjectName => _moduleName.ProjectName;

        //TODO: Move this to the user of the arguments. Event arguments should be pure data.
        public void Navigate(ISelectionService selectionService)
        {
            var selection = new Selection(Exception.LineNumber, Exception.Position, Exception.LineNumber, Exception.Position + Exception.OffendingSymbol.Text.Length - 1);
            selectionService.TrySetActiveSelection(_moduleName, selection);
        }
    }
}
