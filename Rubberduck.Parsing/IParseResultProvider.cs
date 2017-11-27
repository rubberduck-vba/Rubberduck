using System;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing
{
    public class ParseProgressEventArgs : EventArgs
    {
        public ParseProgressEventArgs(QualifiedModuleName module, ParserState state, ParserState oldState)
        {
            Module = module;
            State = state;
            OldState = oldState;
        }

        public QualifiedModuleName Module { get; }

        public ParserState State { get; }

        public ParserState OldState { get; }
    }
}
