using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA.Nodes
{
    public class TypeNode :Node
    {
        private readonly VisualBasic6Parser.TypeStmtContext _context;
        private readonly IdentifierNode _identifier;

        public TypeNode(VisualBasic6Parser.TypeStmtContext context, string scope)
            :base(context, scope, null, new List<Node>())
        {
            _context = context;
            _identifier = new IdentifierNode(_context.ambiguousIdentifier(), scope);

            var children = context.typeStmt_Element();
            foreach (var child in children)
            {
                AddChild(new TypeElementNode(child, scope));
            }
        }

        public IdentifierNode Identifier
        {
            get {return _identifier;}
        }

        public VBAccessibility Accessibility
        {
            get {return (VBAccessibility) Enum.Parse(typeof (VBAccessibility), _context.visibility().GetText());}
        }
    }

    public class TypeElementNode : Node
    {
        private readonly VisualBasic6Parser.TypeStmt_ElementContext _context;
        private readonly IdentifierNode _identifier;

        public TypeElementNode(VisualBasic6Parser.TypeStmt_ElementContext context, string scope)
            :base(context,scope)
        {
            _context = context;
            _identifier = new IdentifierNode(_context.ambiguousIdentifier(), scope);
        }

        public string IdentifierName
        {
            get {return _identifier.Name; }
        }
    }
}
