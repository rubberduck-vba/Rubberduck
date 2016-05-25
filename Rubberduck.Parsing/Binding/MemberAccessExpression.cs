using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Binding
{
    public sealed class MemberAccessExpression : BoundExpression
    {
        private readonly IBoundExpression _lExpression;
        private readonly ParserRuleContext _unrestrictedNameContext;

        public MemberAccessExpression(
            Declaration referencedDeclaration, 
            ExpressionClassification classification, 
            ParserRuleContext context,
            ParserRuleContext unrestrictedNameContext,
            IBoundExpression lExpression)
            : base(referencedDeclaration, classification, context)
        {
            _lExpression = lExpression;
            _unrestrictedNameContext = unrestrictedNameContext;
        }

        public IBoundExpression LExpression
        {
            get
            {
                return _lExpression;
            }
        }

        public ParserRuleContext UnrestrictedNameContext
        {
            get
            {
                return _unrestrictedNameContext;
            }
        }
    }
}
