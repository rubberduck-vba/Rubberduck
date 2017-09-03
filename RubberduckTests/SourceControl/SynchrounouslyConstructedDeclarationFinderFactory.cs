using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.Application;
using System.Collections.Generic;

namespace RubberduckTests.SourceControl
{
    internal class SynchrounouslyConstructedDeclarationFinderFactory : IDeclarationFinderFactory
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