using System.Collections.Generic;
using System.Threading;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA.ReferenceManagement
{
    public interface IReferenceRemover
    {
        void RemoveReferencesBy(QualifiedModuleName module, CancellationToken token);
        void RemoveReferencesBy(IReadOnlyCollection<QualifiedModuleName> modules, CancellationToken token);
        void RemoveReferencesTo(QualifiedModuleName module, CancellationToken token);
        void RemoveReferencesTo(IReadOnlyCollection<QualifiedModuleName> modules, CancellationToken token);
    }
}
