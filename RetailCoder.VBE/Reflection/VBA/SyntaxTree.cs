using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RetailCoderVBE.Reflection.VBA
{
    internal class SyntaxTreeNode
    {
        public SyntaxTreeNode(SyntaxTreeNodeType nodeType, string code, IEnumerable<SyntaxTreeNode> childNodes)
        {
            _nodeType = nodeType;
            _code = code;
            _nodes = childNodes;
        }

        private readonly SyntaxTreeNodeType _nodeType;
        public SyntaxTreeNodeType NodeType { get { return _nodeType; } }

        private readonly string _code;
        public string Code { get { return _code; } }

        private readonly IEnumerable<SyntaxTreeNode> _nodes;
        public IEnumerable<SyntaxTreeNode> ChildNodes { get { return _nodes; } }

        private static readonly IEnumerable<string> _declarationKeywords = new[] 
            {
                "Dim", "Static"
            };

        public SyntaxTreeNode Parse(string code)
        {
            var lines = code.Split('\n').ToList();
            if (lines.Count == 1)
            {
                var line = lines[0];
                if (line.StartsWith("'"))
                {
                    return new SyntaxTreeNode(SyntaxTreeNodeType.Comment, line, new List<SyntaxTreeNode>());
                }
                else
                {
                    if (_declarationKeywords.Any(keyword => line.StartsWith(keyword)))
                    {
                        return new SyntaxTreeNode(SyntaxTreeNodeType.Instruction, line, new List<SyntaxTreeNode>());
                    }
                }
            }

            return null;
        }
    }

    internal enum SyntaxTreeNodeType
    {
        Comment,
        Expression,
        Instruction,
        Branch,
        Loop,
        Jump
    }
}
