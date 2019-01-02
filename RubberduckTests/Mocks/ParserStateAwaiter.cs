using System;
using System.Collections.Generic;
using System.Threading;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Mocks
{
    public class ParserStateAwaiter
    {
        private readonly EventWaitHandle _waitHandle;
        private readonly ICollection<ParserState> _statesToAwait;

        public ParserStateAwaiter(EventWaitHandle waitHandle, ICollection<ParserState> statesToAwait)
        {
            _waitHandle = waitHandle;
            _statesToAwait = statesToAwait;
        }

        public void ParserStateHandler(object requestor, ParserStateEventArgs e)
        {
            if (_statesToAwait.Contains(e.State))
            {
                _waitHandle.Set();
            }
        }
    }
}