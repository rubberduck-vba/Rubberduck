using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Threading;

namespace Rubberduck.Parsing.VBA
{
    public interface IDeclarationResolveRunner
    {
        void ResolveDeclarations(ICollection<QualifiedModuleName> modules, CancellationToken token);
    }
}
