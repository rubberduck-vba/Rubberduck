using Antlr4.Runtime;

namespace Rubberduck.Parsing.Syntax
{
    public abstract class SyntaxNode
    {
        protected SyntaxNode(SyntaxNode parent, IToken token)
            : this(parent, token, new TextSpan(token.Line, token.Column, token.Line, token.Column + token.Text.Length))
        {
        }

        protected SyntaxNode(SyntaxNode parent, IToken token, TextSpan span)
        {
            _parent = parent;
            _token = token;
            _span = span;
        }

        private readonly SyntaxNode _parent;
        public SyntaxNode Parent { get { return _parent; } }

        private readonly TextSpan _span;
        public TextSpan Span { get { return _span; } }

        private readonly IToken _token;
        public IToken Token { get { return _token; } }
    }
}