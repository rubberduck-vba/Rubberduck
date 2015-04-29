using Antlr4.Runtime;

namespace Rubberduck.Parsing.Symbols.Resolver
{
    public interface IResolver<in TContext> where TContext : ParserRuleContext
    {
        Declaration Resolve(TContext identifierContext, QualifiedModuleName currentModule, string currentScope, Declaration qualifier = null);
    }
}