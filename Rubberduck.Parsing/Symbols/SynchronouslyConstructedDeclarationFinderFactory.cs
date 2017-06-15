using Rubberduck.Parsing.Annotations;
using Rubberduck.VBEditor.Application;
using System.Collections.Generic;

namespace Rubberduck.Parsing.Symbols
{
    public class SynchronouslyConstructedDeclarationFinderFactory : IDeclarationFinderFactory 
    {
        public DeclarationFinder Create(IReadOnlyList<Declaration> declarations, IEnumerable<IAnnotation> annotations, IReadOnlyList<UnboundMemberDeclaration> unresolvedMemberDeclarations, IHostApplication hostApp)
        {
            return new DeclarationFinder(declarations, annotations, unresolvedMemberDeclarations, hostApp);
        }
    }
}
