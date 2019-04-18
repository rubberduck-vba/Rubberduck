using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Binding
{
    public sealed class IndexExpression : BoundExpression
    {
        public IndexExpression(
            Declaration referencedDeclaration, 
            ExpressionClassification classification, 
            ParserRuleContext context,
            IBoundExpression lExpression,
            ArgumentList argumentList)
            : base(referencedDeclaration, classification, context)
        {
            LExpression = lExpression;
            ArgumentList = argumentList;
        }

        public IBoundExpression LExpression { get; }
        public ArgumentList ArgumentList { get; }
    }
}
