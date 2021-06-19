using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Refactorings.DeleteDeclarations;

namespace Rubberduck.Refactorings
{
    public interface IEOSContextContentProviderFactory
    {
        IEOSContextContentProvider Create(VBAParser.EndOfStatementContext eosContext, IModuleRewriter rewriter);
    }

    public class EOSContextContentProviderFactory : IEOSContextContentProviderFactory
    {
        public IEOSContextContentProvider Create(VBAParser.EndOfStatementContext eosContext, IModuleRewriter rewriter)
        {
            return new EOSContextContentProvider(eosContext, rewriter);
        }
    }
}
