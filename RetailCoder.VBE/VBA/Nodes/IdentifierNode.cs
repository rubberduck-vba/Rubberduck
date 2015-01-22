using Rubberduck.Extensions;

namespace Rubberduck.VBA.Nodes
{
    public class IdentifierNode : Node
    {
        private readonly VisualBasic6Parser.CertainIdentifierContext _certainContext;
        private readonly VisualBasic6Parser.AmbiguousIdentifierContext _ambiguousContext;

        public IdentifierNode(Selection location, string project, string module, string scope,
            VisualBasic6Parser.CertainIdentifierContext context)
            : base(location, project, module, scope)
        {
            _certainContext = context;
        }

        public IdentifierNode(Selection location, string project, string module, string scope,
            VisualBasic6Parser.AmbiguousIdentifierContext context)
            : base(location, project, module, scope)
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