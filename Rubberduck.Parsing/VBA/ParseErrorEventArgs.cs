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
            _exception = exception;
            _component = component;
        }

        private readonly SyntaxErrorException _exception;
        public SyntaxErrorException Exception { get { return _exception; } }

        private readonly IVBComponent _component;
        public string ComponentName { get { return _component.Name; } }
        public string ProjectName { get { return _component.Collection.Parent.Name; } }

        public void Navigate()
        {
            var selection = new Selection(_exception.LineNumber, _exception.Position, _exception.LineNumber, _exception.Position + _exception.OffendingSymbol.Text.Length - 1);
            using (var module = _component.CodeModule)
            {
                using (var pane = module.CodePane)
                {
                    pane.Selection = selection;
                }
            }
        }
    }
}
