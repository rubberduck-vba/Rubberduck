using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Threading;

namespace Rubberduck.Parsing.VBA
{
    public interface IDeclarationResolveRunner
    {
        void ResolveDeclarations(IReadOnlyCollection<QualifiedModuleName> modules, CancellationToken token);
    }
}
