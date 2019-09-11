using System;
using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Grammar.Abstract.CodePathAnalysis;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Grammar
{
    public partial class VBAPArser
    {
        public partial class IfStmtContext : IBranchNode
        {
            private readonly IDictionary<IExecutionContext, bool> _hasExecuted
                = new Dictionary<IExecutionContext, bool>();

            public bool HasExecuted(IExecutionContext context) => _hasExecuted.TryGetValue(context, out var value) && value;

            public void Execute(IExecutionContext context)
            {
                _hasExecuted[context] = true;
                //if (ConditionExpression.Evaluate<bool>(context))
                //{
                //    // TODO
                //}
            }

            public IEvaluatableNode ConditionExpression { get; set; }
        }
    }

    public partial class VBAParser
    {
        public partial class AnnotationContext
        {
            // todo extend
        }
    }
}
