using Rubberduck.Parsing.Annotations;
using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.Symbols
{
    public class DeclarationFinderFactory : IDeclarationFinderFactory 
    {
        public DeclarationFinder Create(IReadOnlyList<Declaration> declarations, IEnumerable<IAnnotation> annotations, IReadOnlyList<UnboundMemberDeclaration> unresolvedMemberDeclarations, IHostApplication hostApp)
        {
            return new DeclarationFinder(declarations, annotations, unresolvedMemberDeclarations, hostApp);
        }

        public void Release(DeclarationFinder declarationFinder)
        {
        }
    }
}
