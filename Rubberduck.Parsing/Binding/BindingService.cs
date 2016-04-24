using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Parsing.Binding
{
    public sealed class BindingService
    {
        private readonly IBindingContext _defaultBindingContext;
        private readonly IBindingContext _typedBindingContext;
        private readonly IBindingContext _procedurePointerBindingContext;

        public BindingService(
            IBindingContext defaultBindingContext,
            IBindingContext typedBindingContext, 
            IBindingContext procedurePointerBindingContext)
        {
            _defaultBindingContext = defaultBindingContext;
            _typedBindingContext = typedBindingContext;
            _procedurePointerBindingContext = procedurePointerBindingContext;
        }

        public IBoundExpression ResolveDefault(Declaration module, Declaration parent, string expression)
        {
            var expr = Parse(expression);
            return _defaultBindingContext.Resolve(module, parent, expr);
        }

        public IBoundExpression ResolveType(Declaration module, Declaration parent, string expression)
        {
            var expr = Parse(expression);
            return _typedBindingContext.Resolve(module, parent, expr);
        }

        public IBoundExpression ResolveProcedurePointer(Declaration module, Declaration parent, string expression)
        {
            var expr = Parse(expression);
            return _procedurePointerBindingContext.Resolve(module, parent, expr);
        }

        private VBAExpressionParser.ExpressionContext Parse(string expression)
        {
            var stream = new AntlrInputStream(expression);
            var lexer = new VBALexer(stream);
            var tokens = new CommonTokenStream(lexer);
            var parser = new VBAExpressionParser(tokens);
            parser.AddErrorListener(new ExceptionErrorListener());
            var tree = parser.startRule();
            return tree.expression();
        }
    }
}
