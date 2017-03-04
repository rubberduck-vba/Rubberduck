using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Binding
{
    public interface IBoundExpression
    {
        Declaration ReferencedDeclaration { get; }
        ExpressionClassification Classification { get; }
        ParserRuleContext Context { get; }
    }
}
