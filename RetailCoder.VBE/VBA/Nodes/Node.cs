using System.Collections.Generic;
using Rubberduck.Extensions;

namespace Rubberduck.VBA.Nodes
{
    public abstract class Node
    {
        protected Node(Selection location, string project, string module, IList<Node> children = null)
            : this(location, project, module, project + "." + module, children)
        { }

        protected Node(Selection location, string project, string module, string scope, IList<Node> children = null)
        {
            _selection = location;
            _projectName = project;
            _moduleName = module;
            _scope = scope;
            _children = children ?? new List<Node>();
        }

        private readonly string _projectName;
        public string ProjectName { get { return _projectName; } }

        private readonly string _moduleName;
        public string ModuleName { get { return _moduleName; } }

        private readonly string _scope;
        public string Scope { get { return _scope; } }

        private readonly Selection _selection;
        public Selection Selection { get { return _selection; } }

        private readonly IList<Node> _children; 
        public IList<Node> Children { get { return _children; } }
    }

    public class ModuleNode : Node
    {
        public ModuleNode(Selection location, string project, string module, IList<Node> nodes)
            :base(location, project, module, nodes)
        { }
    }
}