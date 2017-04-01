using System;
using System.Collections.Generic;
using System.Threading;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Mocks
{
    internal class ParserStateChangedTestCallbackRunner : IParserStateChangedCallbackRunner
    {
        public void Run(ICollection<Action<CancellationToken>> callbacks, CancellationToken token)
        {
            foreach (var callback in callbacks)
            {
                if (token.IsCancellationRequested) { break; }

                callback(token);
            }
        }
    }
}