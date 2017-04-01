using System;
using System.Collections.Generic;
using System.Threading;

namespace Rubberduck.Parsing.VBA
{
    public interface IParserStateChangedCallbackRunner
    {
        void Run(ICollection<Action<CancellationToken>> callbacks, CancellationToken token);
    }
}