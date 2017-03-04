using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Binding
{
    public sealed class InstanceExpression : BoundExpression
    {
        public InstanceExpression(
            Declaration referencedDeclaration,
            ParserRuleContext context)
            // Note: According to MS VBAL actually a value(?) but we reclassify to bring it more in line with the rest of the binding process.
            : base(referencedDeclaration, ExpressionClassification.Variable, context)
        {
        }
    }
}
