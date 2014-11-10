using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RetailCoderVBE.Reflection.VBA
{
    internal class ModuleNode : SyntaxTreeNode
    {
        public ModuleNode(string projectName, string moduleName, IEnumerable<SyntaxTreeNode> members)
            : base(projectName, null, string.Empty, true)
        {
            _name = moduleName;
            _members = members;
        }

        private readonly string _name;
        public Identifier Identifier
        {
            get
            {
                return new Identifier(Scope, _name, _name);
            }
        }

        private readonly IEnumerable<SyntaxTreeNode> _members;
        public IEnumerable<SyntaxTreeNode> Members { get { return _members; } }
    }
}
