using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA.Nodes
{
    public class IdentifierNode : Node
    {
        private readonly VisualBasic6Parser.CertainIdentifierContext _certainContext;
        private readonly VisualBasic6Parser.AmbiguousIdentifierContext _ambiguousContext;

        public IdentifierNode(VisualBasic6Parser.CertainIdentifierContext context, string scope)
            : base(context, scope)
        {
            _certainContext = context;
        }

        public IdentifierNode(VisualBasic6Parser.AmbiguousIdentifierContext context, string scope)
            : base(context, scope)
        {
            _ambiguousContext = context;
        }

        public string Name
        {
            get
            {
                return _certainContext != null
                    ? _certainContext.IDENTIFIER()[0].GetText()
                    : _ambiguousContext.IDENTIFIER()[0].GetText();
            }
        }
    }
}