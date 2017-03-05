using System;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing
{
    public class ParseProgressEventArgs : EventArgs
    {
        private readonly IVBComponent _component;
        private readonly ParserState _state;
        private readonly ParserState _oldState;

        public ParseProgressEventArgs(IVBComponent component, ParserState state, ParserState oldState)
        {
            _component = component;
            _state = state;
            _oldState = oldState;
        }

        public IVBComponent Component { get { return _component; } }
        public ParserState State { get { return _state; } }
        public ParserState OldState { get { return _oldState; } }
    }
}
