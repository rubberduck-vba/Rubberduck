using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Binding
{
    public sealed class SimpleNameExpression : BoundExpression
    {
        public SimpleNameExpression(Declaration referencedDeclaration, ExpressionClassification classification, ParserRuleContext context)
            : base(referencedDeclaration, classification, context)
        {
        }
    }
}
