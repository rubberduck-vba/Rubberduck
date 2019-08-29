using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Binding
{
    public sealed class IndexExpression : BoundExpression
    {
        public IndexExpression(Declaration referencedDeclaration,
            ExpressionClassification classification,
            ParserRuleContext context,
            IBoundExpression lExpression,
            ArgumentList argumentList,
            bool isArrayAccess = false,
            bool isDefaultMemberAccess = false,
            int defaultMemberRecursionDepth = 0,
            RecursiveDefaultMemberAccessExpression containedDefaultMemberRecursionExpression = null)
            : base(referencedDeclaration, classification, context)
        {
            LExpression = lExpression;
            ArgumentList = argumentList;
            IsArrayAccess = isArrayAccess;
            IsDefaultMemberAccess = isDefaultMemberAccess;
            DefaultMemberRecursionDepth = defaultMemberRecursionDepth;
            ContainedDefaultMemberRecursionExpression = containedDefaultMemberRecursionExpression;
        }

        public IBoundExpression LExpression { get; }
        public ArgumentList ArgumentList { get; }
        public bool IsArrayAccess { get; }
        public bool IsDefaultMemberAccess { get; }
        public int DefaultMemberRecursionDepth { get; }

        public RecursiveDefaultMemberAccessExpression ContainedDefaultMemberRecursionExpression { get; }
    }
}
