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
        public partial class IfStmtContext : IBranchNode
        {
            private readonly IDictionary<IExecutionContext, bool> _hasExecuted
                = new Dictionary<IExecutionContext, bool>();

            public bool HasExecuted(IExecutionContext context) 
                => _hasExecuted.TryGetValue(context, out var value) && value;

            public void Execute(IExecutionContext context) 
                => _hasExecuted[context] = true;

            public IEvaluatableNode ConditionExpression => this.booleanExpression();

            public bool IsReachable { get; set; }
        }

        public partial class SingleLineIfStmtContext : IBranchNode
        {
            private readonly IDictionary<IExecutionContext, bool> _hasExecuted
                = new Dictionary<IExecutionContext, bool>();

            public bool HasExecuted(IExecutionContext context) 
                => _hasExecuted.TryGetValue(context, out var value) && value;

            public void Execute(IExecutionContext context) 
                => _hasExecuted[context] = true;

            public IEvaluatableNode ConditionExpression 
                => this.ifWithEmptyThen()?.booleanExpression() 
                ?? this.ifWithNonEmptyThen()?.booleanExpression();

            public bool IsReachable { get; set; }
        }

        public partial class IfWithEmptyThenContext : IBranchNode
        {
            private readonly IDictionary<IExecutionContext, bool> _hasExecuted
                = new Dictionary<IExecutionContext, bool>();

            public bool HasExecuted(IExecutionContext context) 
                => _hasExecuted.TryGetValue(context, out var value) && value;

            public void Execute(IExecutionContext context) 
                => _hasExecuted[context] = true;

            public IEvaluatableNode ConditionExpression
                => this.booleanExpression();
            public bool IsReachable { get; set; }
        }

        public partial class IfWithNonEmptyThenContext : IBranchNode
        {
            private readonly IDictionary<IExecutionContext, bool> _hasExecuted
                = new Dictionary<IExecutionContext, bool>();
            
            public bool HasExecuted(IExecutionContext context) 
                => _hasExecuted.TryGetValue(context, out var value) && value;

            public void Execute(IExecutionContext context) 
                => _hasExecuted[context] = true;

            public IEvaluatableNode ConditionExpression
                => this.booleanExpression();
            public bool IsReachable { get; set; }
        }

        public partial class ElseIfBlockContext : IBranchNode
        {
            private readonly IDictionary<IExecutionContext, bool> _hasExecuted
                = new Dictionary<IExecutionContext, bool>();
            
            public bool HasExecuted(IExecutionContext context)
                => _hasExecuted.TryGetValue(context, out var value) && value;

            public void Execute(IExecutionContext context) 
                => _hasExecuted[context] = true;

            public IEvaluatableNode ConditionExpression
                => this.booleanExpression();
            public bool IsReachable { get; set; }
        }

        public partial class ElseBlockContext : IBranchNode
        {
            private readonly IDictionary<IExecutionContext, bool> _hasExecuted
                = new Dictionary<IExecutionContext, bool>();
            
            public bool HasExecuted(IExecutionContext context)
                => _hasExecuted.TryGetValue(context, out var value) && value;

            public void Execute(IExecutionContext context) 
                => _hasExecuted[context] = true;

            public IEvaluatableNode ConditionExpression => null;
            public bool IsReachable { get; set; } 
        }
    }
}
