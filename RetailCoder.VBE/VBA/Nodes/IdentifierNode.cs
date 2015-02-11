using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA.Nodes
{
    public class IdentifierNode : Node
    {
        private readonly VisualBasic6Parser.CertainIdentifierContext _certainContext;
        private readonly VisualBasic6Parser.AmbiguousIdentifierContext _ambiguousContext;
        private readonly VisualBasic6Parser.AsTypeClauseContext _asTypeClauseContext;

        public IdentifierNode(VisualBasic6Parser.CertainIdentifierContext context, string scope, VisualBasic6Parser.AsTypeClauseContext asTypeClause = null)
            : base(context, scope)
        {
            _certainContext = context;
            _asTypeClauseContext = asTypeClause;
        }

        public IdentifierNode(VisualBasic6Parser.AmbiguousIdentifierContext context, string scope, VisualBasic6Parser.AsTypeClauseContext asTypeClause = null)
            : base(context, scope)
        {
            _ambiguousContext = context;
            _asTypeClauseContext = asTypeClause;
        }

        public string Name
        {
            get
            {
                return _certainContext != null
                    ? _certainContext.GetText()
                    : _ambiguousContext.GetText();
            }
        }

        public override string ToString()
        {
            return _asTypeClauseContext == null
                ? Name
                : Name + ' ' + _asTypeClauseContext.GetText();
        }
    }
}