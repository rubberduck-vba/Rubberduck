using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;

namespace Rubberduck.VBA
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
            : this(project.Name, nodes)
        {
        }

        public ProjectNode(string projectName, IEnumerable<SyntaxTreeNode> nodes)
            : base(Instruction.Empty(new LogicalCodeLine(projectName, projectName, 0, 0, string.Empty)), string.Empty, null, nodes)
        {
        }
    }
}
