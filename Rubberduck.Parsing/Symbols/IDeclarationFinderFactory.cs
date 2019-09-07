using Rubberduck.Parsing.Annotations;
using System.Collections.Generic;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.Symbols
{
    public interface IDeclarationFinderFactory
    {
        DeclarationFinder Create(
            IReadOnlyList<Declaration> declarations, 
            IEnumerable<IAnnotation> annotations, 
            IReadOnlyList<UnboundMemberDeclaration> unresolvedMemberDeclarations, 
            IReadOnlyDictionary<QualifiedModuleName, IReadOnlyCollection<IdentifierReference>> unboundDefaultMemberAccesses,
            IReadOnlyDictionary<QualifiedModuleName, IReadOnlyCollection<IdentifierReference>> failedLetCoercions,
            IHostApplication hostApp);
        void Release(DeclarationFinder declarationFinder);
    }
}
