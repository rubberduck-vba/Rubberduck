using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Binding
{
    public sealed class MemberAccessExpression : BoundExpression
    {
        public MemberAccessExpression(
            Declaration referencedDeclaration, 
            ExpressionClassification classification, 
            ParserRuleContext context,
            ParserRuleContext unrestrictedNameContext,
            IBoundExpression lExpression)
            : base(referencedDeclaration, classification, context)
        {
            LExpression = lExpression;
            UnrestrictedNameContext = unrestrictedNameContext;
        }

        public IBoundExpression LExpression { get; }
        public ParserRuleContext UnrestrictedNameContext { get; }
    }
}
