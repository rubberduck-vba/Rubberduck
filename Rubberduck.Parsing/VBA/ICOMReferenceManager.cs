using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System.Collections.Generic;
using System.Threading;

namespace Rubberduck.Parsing.VBA
{
    public interface ICOMReferenceManager
    {
        bool LastRunLoadedReferences { get; }
        bool LastRunUnloadedReferences { get; }
        IReadOnlyCollection<ReferencePriorityMap> ProjectReferences { get; }

        void SyncComReferences(IReadOnlyList<IVBProject> projects, CancellationToken token);
    }
}
