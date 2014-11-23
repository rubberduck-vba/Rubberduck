using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;

namespace Rubberduck.VBA.Parser
{
    [ComVisible(false)]
    public class ModuleNode : SyntaxTreeNode
    {
        public ModuleNode(string projectName, string componentName, IEnumerable<SyntaxTreeNode> nodes)
            : base(Instruction.Empty(new LogicalCodeLine(projectName, componentName, 0, 0, string.Empty)), projectName, null, nodes)
        {
        }
    }

    [ComVisible(false)]
    public class ProjectNode : SyntaxTreeNode
    {
        public ProjectNode(VBProject project, IEnumerable<SyntaxTreeNode> nodes)
            : base(Instruction.Empty(new LogicalCodeLine(project.Name, project.Name, 0, 0, string.Empty)), string.Empty, null, nodes)
        {
        }
    }
}
