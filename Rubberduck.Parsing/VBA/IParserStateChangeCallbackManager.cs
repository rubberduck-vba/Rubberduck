using System;
using System.Threading;

namespace Rubberduck.Parsing.VBA
{
    public interface IParserStateChangeCallbackManager
    {
        void RegisterCallback(Action<CancellationToken> callback, ParserState parserState);
        void UnregisterCallback(Action<CancellationToken> callback);
        void RunCallbacks(ParserState state, CancellationToken token);
    }
}