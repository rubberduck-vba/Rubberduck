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
            public void Execute(IExecutionContext context)
            {
                IsReachable = true;
            }

            public bool IsReachable { get; set; }

            public IEvaluatableNode ConditionExpression { get; set; }
        }

        public partial class ForEachStmtContext : ILoopNode
        {
            public void Execute(IExecutionContext context)
            {
                IsReachable = true;
            }

            public bool IsReachable { get; set; }

            public IEvaluatableNode ConditionExpression { get; set; }
        }
        public partial class DoLoopStmtContext : ILoopNode
        {
            public void Execute(IExecutionContext context)
            {
                IsReachable = true;
            }

            public bool IsReachable { get; set; }
            public IEvaluatableNode ConditionExpression { get; set; }
        }
    }
}
