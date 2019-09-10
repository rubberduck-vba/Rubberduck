using System.Collections.Generic;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;

namespace Rubberduck.Inspections.CodePathAnalysis.Nodes
{
    public class ErrorPath : CodePath
    {
        public ErrorPath(IExecutableNode origin) : base(origin) { }
    }

    public class ExecutionContext
    {
        // todo
    }

    public class CodePath
    {
        private readonly ISet<IExecutableNode> _nodes = new HashSet<IExecutableNode>();

        public CodePath(IExecutableNode origin = null)
        {
            Origin = origin;
        }

        public IExecutableNode Origin { get; }

        public void AddNode(IExecutableNode node) => _nodes.Add(node);
    }

    public class BlockNode : NodeBase
    {
        public BlockNode(IParseTree tree) : base(tree) { }
    }

    public class StatementNode : NodeBase, IExecutableNode
    {
        public StatementNode(IParseTree tree) : base(tree)
        {
        }

        /// <summary>
        /// True if node was hit in a code path traversal.
        /// </summary>
        public bool HasExecuted => _hits > 0;

        /// <summary>
        /// The number of times a node was executed in a code path traversal.
        /// </summary>
        public int Hits => _hits;

        private int _hits;
        public virtual void Execute(ExecutionContext context)
        {
            _hits++;
        }
    }
}
