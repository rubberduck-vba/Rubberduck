using System.Collections.Generic;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA.ReferenceManagement;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.VBA.DeclarationCaching
{
    public class ConcurrentlyConstructedDeclarationFinderFactory : IDeclarationFinderFactory
    {
        public DeclarationFinder Create(IReadOnlyList<Declaration> declarations,
            IEnumerable<IParseTreeAnnotation> annotations,
            IReadOnlyDictionary<QualifiedModuleName, LogicalLineStore> logicalLines,
            IReadOnlyDictionary<QualifiedModuleName, IFailedResolutionStore> failedResolutionStores,
            IHostApplication hostApp)
        {
            return new ConcurrentlyConstructedDeclarationFinder(declarations, annotations, logicalLines, failedResolutionStores, hostApp);
        }

        public void Release(DeclarationFinder declarationFinder)
        {
        }
    }
}
