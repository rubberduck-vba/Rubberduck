using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Rubberduck.VBA.Parser
{
    [ComVisible(false)]
    public class ModuleNode : SyntaxTreeNode
    {
        public ModuleNode(string projectName, string componentName, IEnumerable<SyntaxTreeNode> nodes)
            : base(Instruction.Empty(new LogicalCodeLine(projectName, componentName, 0, string.Empty)), projectName, null, nodes)
        {
        }
    }
}
