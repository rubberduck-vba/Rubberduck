using Rubberduck.Parsing.Annotations;
using System.Collections.Generic;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Parsing.VBA.ReferenceManagement;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.Symbols
{
    public interface IDeclarationFinderFactory
    {
        DeclarationFinder Create(IReadOnlyList<Declaration> declarations,
            IEnumerable<IParseTreeAnnotation> annotations,
            IReadOnlyDictionary<QualifiedModuleName, LogicalLineStore> logicalLines,
            IReadOnlyDictionary<QualifiedModuleName, IFailedResolutionStore> failedResolutionStores,
            IHostApplication hostApp);
        void Release(DeclarationFinder declarationFinder);
    }
}
