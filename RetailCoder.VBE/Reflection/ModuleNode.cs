using Rubberduck.Reflection.VBA;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Reflection
{
    internal class ModuleNode : SyntaxTreeNode
    {
        public ModuleNode(string projectName, string componentName, IEnumerable<SyntaxTreeNode> nodes)
            : base(Instruction.Empty(new LogicalCodeLine(projectName, componentName, 0, string.Empty)), projectName, null, true)
        {
            _nodes = nodes;
        }

        private readonly IEnumerable<SyntaxTreeNode> _nodes;
        public IEnumerable<SyntaxTreeNode> Nodes { get { return _nodes; } }
    }
}
