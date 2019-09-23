using Rubberduck.Parsing.Grammar.Abstract.CodePathAnalysis;
using System;

namespace Rubberduck.CodeAnalysis.CodePathAnalysis.Execution.ExtendedNodeVisitor
{
    public class MaxIterationsReachedException : Exception
    {
        public MaxIterationsReachedException(IExtendedNode node)
            : base("Maximum number of iterations was reached; code path is possibly in an infinite loop.")
        {
            Node = node;
        }

        public IExtendedNode Node { get; }
    }
}
