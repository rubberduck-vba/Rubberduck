using Rubberduck.Parsing.Grammar.Abstract.CodePathAnalysis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Parsing.Grammar
{
    public partial class VBAParser
    {
        public partial class BooleanExpressionContext : IEvaluatableNode
        {
            public T Evaluate<T>(IExecutionContext context)
            {
                if (typeof(T) != typeof(bool))
                {
                    throw new NotSupportedException();
                }

                IsReachable = true;
                return default;
            }

            public bool IsReachable { get; set; }
        }

        public partial class SimpleNameExprContext : IEvaluatableNode
        {
            public T Evaluate<T>(IExecutionContext context)
            {
                if (typeof(T) != typeof(bool))
                {
                    throw new NotSupportedException();
                }

                IsReachable = true;
                return default;
            }

            public bool IsReachable { get; set; }
        }
    }
}
