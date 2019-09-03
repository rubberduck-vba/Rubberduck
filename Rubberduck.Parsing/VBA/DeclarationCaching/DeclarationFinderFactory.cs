using System.Collections.Generic;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.VBA.DeclarationCaching
{
    public class DeclarationFinderFactory : IDeclarationFinderFactory 
    {
        public DeclarationFinder Create(IReadOnlyList<Declaration> declarations, 
            IEnumerable<IParseTreeAnnotation> annotations, 
            IReadOnlyList<UnboundMemberDeclaration> unresolvedMemberDeclarations, 
            IReadOnlyDictionary<QualifiedModuleName, IReadOnlyCollection<IdentifierReference>> unboundDefaultMemberAccesses,
            IHostApplication hostApp)
        {
            return new DeclarationFinder(declarations, annotations, unresolvedMemberDeclarations, unboundDefaultMemberAccesses, hostApp);
        }

        public void Release(DeclarationFinder declarationFinder)
        {
        }
    }
}
