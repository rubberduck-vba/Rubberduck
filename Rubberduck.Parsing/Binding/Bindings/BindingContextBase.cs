using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using NLog;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Binding
{
    public abstract class BindingContextBase : IBindingContext
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();



        protected IExpressionBinding HandleUnexpectedExpressionType(ParserRuleContext expression)
        {
            Logger.Warn($"Unexpected context type {expression.GetType()}");
            return new FailedExpressionBinding(expression);
        }

        public abstract IBoundExpression Resolve(Declaration module,
            Declaration parent,
            ParserRuleContext expression,
            IBoundExpression withBlockVariable,
            StatementResolutionContext statementContext,
            bool requiresLetCoercion = false,
            bool isLetAssignment = false);

        public abstract IExpressionBinding BuildTree(Declaration module,
            Declaration parent,
            ParserRuleContext expression,
            IBoundExpression withBlockVariable,
            StatementResolutionContext statementContext,
            bool requiresLetCoercion = false,
            bool isLetAssignment = false);
    }
}