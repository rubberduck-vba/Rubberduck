using System.Collections.Generic;
using System.Threading;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA.DeclarationResolving
{
    public interface IDeclarationResolveRunner
    {
        void ResolveDeclarations(IReadOnlyCollection<QualifiedModuleName> modules, CancellationToken token);
    }
}
