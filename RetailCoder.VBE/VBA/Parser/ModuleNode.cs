using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;

namespace Rubberduck.VBA.Parser
{
    [ComVisible(false)]
    public class ModuleNode : SyntaxTreeNode
    {
        public ModuleNode(string projectName, string componentName, IEnumerable<SyntaxTreeNode> nodes, bool isClassModule)
            : base(Instruction.Empty(new LogicalCodeLine(projectName, componentName, 0, 0, string.Empty)), projectName, null, nodes)
        {
            _isClassModule = isClassModule;
            _projectName = projectName;
            _componentName = componentName;
        }

        public Identifier Identifier { get {  return new Identifier(_projectName, _componentName, _componentName);} }

        private readonly bool _isClassModule;
        private readonly string _projectName;
        private readonly string _componentName;
        public bool IsClassModule { get { return _isClassModule; } }
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
