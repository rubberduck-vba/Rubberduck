using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.Parsing.Grammar.Abstract.CodePathAnalysis;

namespace Rubberduck.Parsing.Grammar
{
    public partial class VBAParser
    {
        public partial class ForNextStmtContext : ILoopNode
        {
            private readonly IDictionary<IExecutionContext, bool> _hasExecuted
                = new Dictionary<IExecutionContext, bool>();

            public bool HasExecuted(IExecutionContext context) 
                => _hasExecuted.TryGetValue(context, out var value) && value;

            public void Execute(IExecutionContext context) 
                => _hasExecuted[context] = true;

            public IEvaluatableNode ConditionExpression { get; set; }
        }

        public partial class ForEachStmtContext : ILoopNode
        {
            private readonly IDictionary<IExecutionContext, bool> _hasExecuted
                = new Dictionary<IExecutionContext, bool>();

            public bool HasExecuted(IExecutionContext context) 
                => _hasExecuted.TryGetValue(context, out var value) && value;

            public void Execute(IExecutionContext context) 
                => _hasExecuted[context] = true;

            public IEvaluatableNode ConditionExpression { get; set; }
        }
        public partial class DoLoopStmtContext : ILoopNode
        {
            private readonly IDictionary<IExecutionContext, bool> _hasExecuted
                = new Dictionary<IExecutionContext, bool>();

            public bool HasExecuted(IExecutionContext context) 
                => _hasExecuted.TryGetValue(context, out var value) && value;

            public void Execute(IExecutionContext context) 
                => _hasExecuted[context] = true;

            public IEvaluatableNode ConditionExpression { get; set; }
        }
    }
}
