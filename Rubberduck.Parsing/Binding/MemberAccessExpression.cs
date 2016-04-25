using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Binding
{
    public sealed class MemberAccessExpression : BoundExpression
    {
        private readonly IBoundExpression _lExpression;

        public MemberAccessExpression(
            Declaration referencedDeclaration, 
            ExpressionClassification classification, 
            ParserRuleContext context,
            IBoundExpression lExpression)
            : base(referencedDeclaration, classification, context)
        {
            _lExpression = lExpression;
        }

        public IBoundExpression LExpression
        {
            get
            {
                return _lExpression;
            }
        }
    }
}
