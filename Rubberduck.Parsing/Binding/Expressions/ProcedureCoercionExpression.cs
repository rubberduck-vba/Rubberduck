using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Binding
{
    public class ProcedureCoercionExpression : BoundExpression
    {
        public ProcedureCoercionExpression(
            Declaration referencedDeclaration,
            ExpressionClassification classification,
            ParserRuleContext context,
            IBoundExpression wrappedExpression)
            : base(referencedDeclaration, classification, context)
        {
            WrappedExpression = wrappedExpression;

            //This works around a problem with the ordering of references between array accesses on (recursive) default member accesses
            //and from subsequent procedure coercion. 
            DefaultMemberRecursionDepth = wrappedExpression is IndexExpression indexExpression
                ? indexExpression.DefaultMemberRecursionDepth + 1
                : 1;
        }

        public IBoundExpression WrappedExpression { get; }
        public int DefaultMemberRecursionDepth { get; }
    }
}