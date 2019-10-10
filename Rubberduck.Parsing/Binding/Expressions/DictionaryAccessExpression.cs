using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Binding
{
    public sealed class DictionaryAccessExpression : BoundExpression
    {
        public DictionaryAccessExpression(
            Declaration referencedDeclaration,
            ExpressionClassification classification,
            ParserRuleContext context,
            IBoundExpression lExpression,
            ArgumentList argumentList,
            ParserRuleContext defaultMemberContext,
            int defaultMemberRecursionDepth = 1,
            RecursiveDefaultMemberAccessExpression containedDefaultMemberRecursionExpression = null)
            : base(referencedDeclaration, classification, context)
        {
            LExpression = lExpression;
            ArgumentList = argumentList;
            DefaultMemberRecursionDepth = defaultMemberRecursionDepth;
            ContainedDefaultMemberRecursionExpression = containedDefaultMemberRecursionExpression;
            DefaultMemberContext = defaultMemberContext;
        }

        public IBoundExpression LExpression { get; }
        public ArgumentList ArgumentList { get; }
        public int DefaultMemberRecursionDepth { get; }
        public ParserRuleContext DefaultMemberContext { get; }
        public RecursiveDefaultMemberAccessExpression ContainedDefaultMemberRecursionExpression { get; }
    }
}