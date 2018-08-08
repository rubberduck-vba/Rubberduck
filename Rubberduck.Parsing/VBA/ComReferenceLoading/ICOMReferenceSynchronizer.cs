using System.Collections.Generic;
using System.Threading;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.VBA.ComReferenceLoading
{
    public interface ICOMReferenceSynchronizer
    {
        bool LastSyncOfCOMReferencesLoadedReferences { get; }
        IEnumerable<QualifiedModuleName> COMReferencesUnloadedUnloadedInLastSync { get; }

        void SyncComReferences(IReadOnlyList<IVBProject> projects, CancellationToken token);
    }
}
