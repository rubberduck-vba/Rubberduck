using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace Rubberduck.Parsing.VBA
{
    public class ParserStateChangedCallbackRunner : IParserStateChangedCallbackRunner
    {
        public void Run(ICollection<Action<CancellationToken>> callbacks, CancellationToken token)
        {
            foreach (var callback in callbacks)
            {
                if (token.IsCancellationRequested) { break; }

                Task.Run(() => callback(token), token);
            }
        }
    }
}