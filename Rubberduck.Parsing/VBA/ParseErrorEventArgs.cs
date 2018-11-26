using System;
using Rubberduck.Parsing.VBA.Parsing.ParsingExceptions;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

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

        public void Navigate(IVBE vbe)
        {
            var selection = new Selection(Exception.LineNumber, Exception.Position, Exception.LineNumber, Exception.Position + Exception.OffendingSymbol.Text.Length - 1);
            
            if (!_moduleName.TryGetComponent(vbe, out var component))
            {
                return;
            }

            using (component)
            using (var module = component.CodeModule)
            using (var pane = module.CodePane)
            {
                pane.Selection = selection;
            }
        }
    }
}
