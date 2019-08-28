using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Binding
{
    public class RecursiveDefaultMemberAccessExpression : BoundExpression
    {
        public RecursiveDefaultMemberAccessExpression(
            Declaration referencedDeclaration,
            ExpressionClassification classification,
            ParserRuleContext context,
            int defaultMemberRecursionDepth = 0,
            RecursiveDefaultMemberAccessExpression containedDefaultMemberRecursionExpression = null)
            : base(referencedDeclaration, classification, context)
        {
            DefaultMemberRecursionDepth = defaultMemberRecursionDepth;
            ContainedDefaultMemberRecursionExpression = containedDefaultMemberRecursionExpression;
        }
        
        public int DefaultMemberRecursionDepth { get; }
        public RecursiveDefaultMemberAccessExpression ContainedDefaultMemberRecursionExpression { get; }
    }
}