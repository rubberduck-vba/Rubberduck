using Rubberduck.Parsing;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA.Nodes
{
    public class IdentifierNode : Node
    {
        private readonly VBAParser.CertainIdentifierContext _certainContext;
        private readonly VBAParser.AmbiguousIdentifierContext _ambiguousContext;
        private readonly VBAParser.AsTypeClauseContext _asTypeClauseContext;

        public IdentifierNode(VBAParser.CertainIdentifierContext context, string scope, VBAParser.AsTypeClauseContext asTypeClause = null)
            : base(context, scope)
        {
            _certainContext = context;
            _asTypeClauseContext = asTypeClause;
        }

        public IdentifierNode(VBAParser.AmbiguousIdentifierContext context, string scope, VBAParser.AsTypeClauseContext asTypeClause = null)
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