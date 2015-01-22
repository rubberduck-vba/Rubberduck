using System;
using Rubberduck.Extensions;

namespace Rubberduck.VBA.Nodes
{
    public class EnumNode : Node
    {
        private readonly VisualBasic6Parser.EnumerationStmtContext _context;
        private readonly IdentifierNode _identifier;

        public EnumNode(Selection location, string project, string module, string scope,
            VisualBasic6Parser.EnumerationStmtContext context)
            :base(location, project, module, scope)
        {
            _context = context;
            _identifier = new IdentifierNode(location, project, module, scope, _context.ambiguousIdentifier());

            var children = context.enumerationStmt_Constant();
            foreach (var child in children)
            {
                Children.Add(new EnumConstNode(location, project, module, scope, child));
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

        public EnumConstNode(Selection location, string project, string module, string scope,
            VisualBasic6Parser.EnumerationStmt_ConstantContext context)
            :base(location, project, module, scope)
        {
            _context = context;
            _identifier = new IdentifierNode(location, project, module, scope, _context.ambiguousIdentifier());
        }

        public string SpecifiedValue
        {
            get { return _context.valueStmt().GetText(); }
        }
    }
}