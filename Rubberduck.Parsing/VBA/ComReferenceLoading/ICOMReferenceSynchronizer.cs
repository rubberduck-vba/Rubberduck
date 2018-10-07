using System.Collections.Generic;
using System.Threading;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA.ComReferenceLoading
{
    public interface ICOMReferenceSynchronizer
    {
        bool LastSyncOfCOMReferencesLoadedReferences { get; }
        IEnumerable<QualifiedModuleName> COMReferencesUnloadedInLastSync { get; }
        IEnumerable<(string projectId, string referencedProjectId)> COMReferencesAffectedByPriorityChangesInLastSync { get; }

        void SyncComReferences(CancellationToken token);
    }
}
