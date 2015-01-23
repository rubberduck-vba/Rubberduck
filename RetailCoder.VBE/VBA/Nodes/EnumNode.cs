using System;
using System.Collections.Generic;
using Rubberduck.Extensions;

namespace Rubberduck.VBA.Nodes
{
    public class EnumNode : Node
    {
        private readonly VisualBasic6Parser.EnumerationStmtContext _context;
        private readonly IdentifierNode _identifier;

        public EnumNode(VisualBasic6Parser.EnumerationStmtContext context, string scope)
            :base(context, scope, null, new List<Node>())
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
            get { return (VBAccessibility) Enum.Parse(typeof (VBAccessibility), _context.visibility().GetText()); }
        }
    }

    public class EnumConstNode : Node
    {
        private readonly VisualBasic6Parser.EnumerationStmt_ConstantContext _context;
        private readonly IdentifierNode _identifier;

        public EnumConstNode(VisualBasic6Parser.EnumerationStmt_ConstantContext context, string scope)
            :base(context, scope)
        {
            _context = context;
            _identifier = new IdentifierNode(_context.ambiguousIdentifier(), scope);
        }

        public string SpecifiedValue
        {
            get { return _context.valueStmt().GetText(); }
        }
    }
}