using System;
using System.Threading;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing
{
    public class ParseProgressEventArgs : EventArgs
    {
        public ParseProgressEventArgs(QualifiedModuleName module, ParserState state, ParserState oldState, CancellationToken token)
        {
            Module = module;
            State = state;
            OldState = oldState;
            Token = token;
        }

        public QualifiedModuleName Module { get; }
        public ParserState State { get; }
        public ParserState OldState { get; }
        public CancellationToken Token { get; }
    }
}
