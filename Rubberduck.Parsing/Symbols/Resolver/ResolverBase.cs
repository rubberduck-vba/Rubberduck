using Antlr4.Runtime;

namespace Rubberduck.Parsing.Symbols.Resolver
{
    public abstract class ResolverBase<TContext> : IResolver<TContext> 
        where TContext : ParserRuleContext
    {
        protected ResolverBase(Declarations declarations)
        {
            _declarations = declarations;
        }

        private readonly Declarations _declarations;
        protected Declarations Declarations { get { return _declarations; } }

        public abstract Declaration Resolve(TContext identifierContext, QualifiedModuleName currentModule, string currentScope, Declaration qualifier = null);
    }
}