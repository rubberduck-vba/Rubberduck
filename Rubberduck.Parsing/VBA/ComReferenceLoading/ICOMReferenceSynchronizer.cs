using System.Collections.Generic;
using System.Threading;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA.ComReferenceLoading
{
    public interface ICOMReferenceSynchronizer
    {
        bool LastSyncOfCOMReferencesLoadedReferences { get; }
        IEnumerable<QualifiedModuleName> COMReferencesUnloadedUnloadedInLastSync { get; }

        void SyncComReferences(CancellationToken token);
    }
}
