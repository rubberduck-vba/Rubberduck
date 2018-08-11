using System.Collections.Generic;
using System.Threading;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA.ReferenceManagement
{
    public interface IReferenceResolveRunner
    {
        void ResolveReferences(IReadOnlyCollection<QualifiedModuleName> toResolve, CancellationToken token);
    }
}
