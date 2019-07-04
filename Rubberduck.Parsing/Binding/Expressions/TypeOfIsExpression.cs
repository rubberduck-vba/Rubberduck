using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Binding
{
    public sealed class TypeOfIsExpression : BoundExpression
    {
        public TypeOfIsExpression(
            Declaration referencedDeclaration,
            ParserRuleContext context,
            IBoundExpression expression,
            IBoundExpression typeExpression)
            : base(referencedDeclaration, ExpressionClassification.Value, context)
        {
            Expression = expression;
            TypeExpression = typeExpression;
        }

        public IBoundExpression Expression { get; }
        public IBoundExpression TypeExpression { get; }
    }
}
