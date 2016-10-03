using System;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.DisposableWrappers;

namespace Rubberduck.Parsing
{
    public class ParseProgressEventArgs : EventArgs
    {
        private readonly VBComponent _component;
        private readonly ParserState _state;
        private readonly ParserState _oldState;

        public ParseProgressEventArgs(VBComponent component, ParserState state, ParserState oldState)
        {
            _component = component;
            _state = state;
            _oldState = oldState;
        }

        public VBComponent Component { get { return _component; } }
        public ParserState State { get { return _state; } }
        public ParserState OldState { get { return _oldState; } }
    }
}
