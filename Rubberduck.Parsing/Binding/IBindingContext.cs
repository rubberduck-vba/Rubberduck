using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Binding
{
    public interface IBindingContext
    {
        IBoundExpression Resolve(Declaration module, Declaration parent, IParseTree expression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext, bool requiresLetCoercion = false, bool isLetAssignment = false);
        IExpressionBinding BuildTree(Declaration module, Declaration parent, IParseTree expression, IBoundExpression withBlockVariable, StatementResolutionContext statementContext, bool requiresLetCoercion = false, bool isLetAssignment = false);
    }
}
