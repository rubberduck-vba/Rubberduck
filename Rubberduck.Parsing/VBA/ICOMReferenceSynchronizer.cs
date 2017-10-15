﻿using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System.Collections.Generic;
using System.Threading;

namespace Rubberduck.Parsing.VBA
{
    public interface ICOMReferenceSynchronizer
    {
        bool LastSyncOfCOMReferencesLoadedReferences { get; }
        IEnumerable<QualifiedModuleName> COMReferencesUnloadedUnloadedInLastSync { get; }

        void SyncComReferences(IReadOnlyList<IVBProject> projects, CancellationToken token);
    }
}
