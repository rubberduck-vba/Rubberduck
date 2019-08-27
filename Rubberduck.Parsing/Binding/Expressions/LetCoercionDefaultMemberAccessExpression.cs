using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Binding
{
    public class LetCoercionDefaultMemberAccessExpression : BoundExpression
    {
        public LetCoercionDefaultMemberAccessExpression(
            Declaration referencedDeclaration,
            ExpressionClassification classification,
            ParserRuleContext context,
            IBoundExpression wrappedExpression,
            int defaultMemberRecursionDepth = 0,
            RecursiveDefaultMemberAccessExpression containedDefaultMemberRecursionExpression = null)
            : base(referencedDeclaration, classification, context)
        {
            WrappedExpression = wrappedExpression;
            DefaultMemberRecursionDepth = defaultMemberRecursionDepth;
            ContainedDefaultMemberRecursionExpression = containedDefaultMemberRecursionExpression;
        }

        public IBoundExpression WrappedExpression { get; }
        public int DefaultMemberRecursionDepth { get; }
        public RecursiveDefaultMemberAccessExpression ContainedDefaultMemberRecursionExpression { get; }
    }
}