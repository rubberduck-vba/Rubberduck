using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Binding
{
    public abstract class BoundExpression : IBoundExpression
    {
        protected BoundExpression(Declaration referencedDeclaration, ExpressionClassification classification, ParserRuleContext context)
        {
            ReferencedDeclaration = referencedDeclaration;
            Classification = classification;
            Context = context;
        }

        public ExpressionClassification Classification { get; }

        public Declaration ReferencedDeclaration { get; }

        public ParserRuleContext Context { get; }
    }
}
