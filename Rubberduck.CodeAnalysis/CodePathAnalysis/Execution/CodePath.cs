using Rubberduck.Parsing.Grammar.Abstract.CodePathAnalysis;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.CodeAnalysis.CodePathAnalysis.Execution
{
    public class CodePath
    {
        private readonly List<IExtendedNode> _nodes;

        public CodePath(IEnumerable<IExtendedNode> nodes = null, bool isErrorPath = false)
        {
            _nodes = new List<IExtendedNode>(nodes ?? Enumerable.Empty<IExtendedNode>());
            IsErrorPath = isErrorPath;
        }

        public bool IsErrorPath { get; }

        public IExtendedNode this[int index] => _nodes[index];

        public int Count => _nodes.Count;

        internal void Add(IExtendedNode node) => _nodes.Add(node);

        internal void AddRange(IEnumerable<IExtendedNode> nodes) 
        {
            foreach (var node in nodes)
            {
                Add(node);
            }
        }
    }
}
