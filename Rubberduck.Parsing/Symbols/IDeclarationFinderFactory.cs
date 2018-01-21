using System;
using Rubberduck.Parsing.Annotations;
using Rubberduck.VBEditor.Application;
using System.Collections.Generic;

namespace Rubberduck.Parsing.Symbols
{
    public interface IDeclarationFinderFactory
    {
        DeclarationFinder Create(IReadOnlyList<Declaration> declarations, IEnumerable<IAnnotation> annotations, IReadOnlyList<UnboundMemberDeclaration> unresolvedMemberDeclarations, IHostApplication hostApp);
        void Release(DeclarationFinder declarationFinder);
    }
}
