using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Rubberduck.VBA.Parser;
using Rubberduck.VBA.Parser.Grammar;

namespace Rubberduck.Reflection
{
    internal class ModuleNode : SyntaxTreeNode
    {
        public ModuleNode(string projectName, string componentName, IEnumerable<SyntaxTreeNode> nodes)
            : base(Instruction.Empty(new LogicalCodeLine(projectName, componentName, 0, string.Empty)), projectName, null, nodes)
        {
        }
    }
}
