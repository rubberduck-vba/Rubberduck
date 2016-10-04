using System;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.DisposableWrappers;
using Rubberduck.VBEditor.DisposableWrappers.VBA;

namespace Rubberduck.Parsing.VBA
{
    public class ParseErrorEventArgs : EventArgs
    {
        public ParseErrorEventArgs(SyntaxErrorException exception, VBComponent component)
        {
            _exception = exception;
            _component = component;
        }

        private readonly SyntaxErrorException _exception;
        public SyntaxErrorException Exception { get { return _exception; } }

        private readonly VBComponent _component;
        public string ComponentName { get { return _component.Name; } }
        public string ProjectName { get { return _component.Collection.Parent.Name; } }

        public void Navigate()
        {
            var selection = new Selection(_exception.LineNumber, _exception.Position, _exception.LineNumber, _exception.Position + _exception.OffendingSymbol.Text.Length - 1);
            using (var module = _component.CodeModule)
            using (var pane = module.CodePane)
            {
                pane.SetSelection(selection);
            }
        }
    }
}
