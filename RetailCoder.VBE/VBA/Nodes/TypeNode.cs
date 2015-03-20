using System;
using System.Collections.Generic;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA.Nodes
{
    public class TypeNode : Node
    {
        private readonly VBAParser.TypeStmtContext _context;
        private readonly IdentifierNode _identifier;

        public TypeNode(VBAParser.TypeStmtContext context, string scope)
            : base(context, scope, null, new List<Node>())
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
            get { return _identifier; }
        }

        public Accessibility Accessibility
        {
            get { return (Accessibility)Enum.Parse(typeof(Accessibility), _context.visibility().GetText()); }
        }
    }

    public class TypeElementNode : Node
    {
        private readonly VBAParser.TypeStmt_ElementContext _context;
        private readonly IdentifierNode _identifier;

        public TypeElementNode(VBAParser.TypeStmt_ElementContext context, string scope)
            : base(context, scope)
        {
            _context = context;
            _identifier = new IdentifierNode(_context.ambiguousIdentifier(), scope);
        }

        public string IdentifierName
        {
            get { return _identifier.Name; }
        }
    }
}