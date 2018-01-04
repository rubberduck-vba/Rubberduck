using System;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.Symbols.ParsingExceptions;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.VBA
{
    public class ParseErrorEventArgs : EventArgs
    {
        public ParseErrorEventArgs(SyntaxErrorException exception, IVBComponent component)
        {
            Exception = exception;
            _component = component;
        }

        public SyntaxErrorException Exception { get; }

        private readonly IVBComponent _component;
        public string ComponentName => _component.Name;
        public string ProjectName => _component.Collection.Parent.Name;

        public void Navigate()
        {
            var selection = new Selection(Exception.LineNumber, Exception.Position, Exception.LineNumber, Exception.Position + Exception.OffendingSymbol.Text.Length - 1);
            var module = _component.CodeModule;
            var pane = module.CodePane;
            {
                pane.Selection = selection;
            }
        }
    }
}
