using Rubberduck.Parsing.Grammar.Abstract.CodePathAnalysis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Parsing.Grammar.PartialExtensions
{
    public partial class VBAParser
    {
        public partial class BlockStmtContext : IExecutableNode
        {
            private readonly IDictionary<IExecutionContext, bool> _hasExecuted
                = new Dictionary<IExecutionContext, bool>();

            public bool HasExecuted(IExecutionContext context) 
                => _hasExecuted.TryGetValue(context, out var value) && value;

            public void Execute(IExecutionContext context)
                => _hasExecuted[context] = true;
        }
    }
}
