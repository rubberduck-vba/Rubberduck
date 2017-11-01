using System;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing
{
    public class ParseProgressEventArgs : EventArgs
    {
        private readonly QualifiedModuleName _module;
        private readonly ParserState _state;
        private readonly ParserState _oldState;

        public ParseProgressEventArgs(QualifiedModuleName module, ParserState state, ParserState oldState)
        {
            _module = module;
            _state = state;
            _oldState = oldState;
        }

        public QualifiedModuleName Module { get { return _module; } }
        public ParserState State { get { return _state; } }
        public ParserState OldState { get { return _oldState; } }
    }
}
