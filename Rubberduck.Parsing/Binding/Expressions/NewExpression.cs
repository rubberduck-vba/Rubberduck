using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Binding
{
    public sealed class NewExpression : BoundExpression
    {
        private readonly IBoundExpression _typeExpression;

        public NewExpression(
            Declaration referencedDeclaration,
            ParserRuleContext context,
            IBoundExpression typeExpression)
            // Marked as Variable instead of Value to integrate into rest of binding process.
            : base(referencedDeclaration, ExpressionClassification.Variable, context)
        {
            _typeExpression = typeExpression;
        }

        public IBoundExpression TypeExpression
        {
            get
            {
                return _typeExpression;
            }
        }
    }
}
