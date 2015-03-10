using System;
using System.Collections.Generic;
using Rubberduck.Parsing;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA.Nodes
{
    public class EnumNode : Node
    {
        private readonly VBAParser.EnumerationStmtContext _context;
        private readonly IdentifierNode _identifier;

        public EnumNode(VBAParser.EnumerationStmtContext context, string scope)
            : base(context, scope, null, new List<Node>())
        {
            _context = context;
            _identifier = new IdentifierNode(_context.ambiguousIdentifier(), scope);

            var children = context.enumerationStmt_Constant();
            foreach (var child in children)
            {
                AddChild(new EnumConstNode(child, scope));
            }
        }

        public IdentifierNode Identifier
        {
            get { return _identifier; }
        }

        public VBAccessibility Accessibility
        {
            get { return (VBAccessibility)Enum.Parse(typeof(VBAccessibility), _context.visibility().GetText()); }
        }
    }

    public class EnumConstNode : Node
    {
        private readonly VBAParser.EnumerationStmt_ConstantContext _context;
        private readonly IdentifierNode _identifier;

        public EnumConstNode(VBAParser.EnumerationStmt_ConstantContext context, string scope)
            : base(context, scope)
        {
            _context = context;
            _identifier = new IdentifierNode(_context.ambiguousIdentifier(), scope);
        }

        public string IdentifierName
        {
            get { return _identifier.Name; }
        }

        public string SpecifiedValue
        {
            get { return _context.valueStmt().GetText(); }
        }
    }
}
