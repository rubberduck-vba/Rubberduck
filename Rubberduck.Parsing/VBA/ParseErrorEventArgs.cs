using System;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEInterfaces;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.Parsing.VBA
{
    public class ParseErrorEventArgs : EventArgs
    {
        public ParseErrorEventArgs(SyntaxErrorException exception, VBComponent component, IRubberduckFactory<IRubberduckCodePane> factory)
        {
            _exception = exception;
            _component = component;
            _factory = factory;
        }

        private readonly SyntaxErrorException _exception;
        private readonly IRubberduckFactory<IRubberduckCodePane> _factory;
        public SyntaxErrorException Exception { get { return _exception; } }

        private readonly VBComponent _component;
        public string ComponentName { get { return _component.Name; } }
        public string ProjectName { get { return _component.Collection.Parent.Name; } }

        public void Navigate()
        {
            var selection = new Selection(_exception.LineNumber, _exception.Position, _exception.LineNumber, _exception.Position + _exception.OffendingSymbol.Text.Length - 1);
            var codePane = _factory.Create(_component.CodeModule.CodePane);
            codePane.Selection = selection;
        }
    }
}