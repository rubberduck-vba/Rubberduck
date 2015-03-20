using System.Collections.Generic;
using Antlr4.Runtime;

namespace Rubberduck.VBA.Nodes
{
    public class ModuleNode : Node
    {
        /// <param name="context">The parser rule context, obtained from an ANTLR-generated parser method.</param>
        /// <param name="project">The name of the VBA project the code belongs to.</param>
        /// <param name="component">The name of the VBA component (Module) the code belongs to.</param>
        public ModuleNode(ParserRuleContext context, string project, string component, ICollection<Node> children)
            : base(context, project, component, children)
        {
            _project = project;
            _component = component;
        }

        private readonly string _project;

        /// <summary>
        /// The name of the VBA project the code belongs to.
        /// </summary>
        public string ProjectName { get { return _project; } }

        private readonly string _component;

        /// <summary>
        /// The name of the VBA component (Module) the code belongs to.
        /// </summary>
        public string ComponentName { get { return _component; } }
    }
}