using Rubberduck.Parsing.Rewriter;
using Rubberduck.Refactorings.DeleteDeclarations;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings
{
    public interface IDeleteDeclarationEndOfStatementContentModifierFactory
    {
        IDeleteDeclarationEndOfStatementContentModifier Create();
    }

    public interface IDeleteDeclarationEndOfStatementContentModifier
    {
        void ModifyEndOfStatementContextContent(IDeclarationDeletionTarget deleteTarget, IDeleteDeclarationModifyEndOfStatementContentModel model, IModuleRewriter rewriter);
    }

    public class DeleteDeclarationEndOfStatementContentModifierFactory : IDeleteDeclarationEndOfStatementContentModifierFactory
    {
        private readonly IEOSContextContentProviderFactory _contentProviderFactory;
        public DeleteDeclarationEndOfStatementContentModifierFactory(IEOSContextContentProviderFactory contentProviderFactory)
        {
            _contentProviderFactory = contentProviderFactory;
        }

        public IDeleteDeclarationEndOfStatementContentModifier Create()
        {
            return new DeleteDeclarationEndOfStatementContentModifier(_contentProviderFactory);
        }
    }
}
