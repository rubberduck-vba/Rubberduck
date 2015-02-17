using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA.Nodes
{
    public class IdentifierNode : Node
    {
        private readonly VBParser.CertainIdentifierContext _certainContext;
        private readonly VBParser.AmbiguousIdentifierContext _ambiguousContext;
        private readonly VBParser.AsTypeClauseContext _asTypeClauseContext;

        public IdentifierNode(VBParser.CertainIdentifierContext context, string scope, VBParser.AsTypeClauseContext asTypeClause = null)
            : base(context, scope)
        {
            _certainContext = context;
            _asTypeClauseContext = asTypeClause;
        }

        public IdentifierNode(VBParser.AmbiguousIdentifierContext context, string scope, VBParser.AsTypeClauseContext asTypeClause = null)
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