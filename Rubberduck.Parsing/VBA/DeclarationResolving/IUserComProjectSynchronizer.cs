using System.Collections.Generic;

namespace Rubberduck.Parsing.VBA.DeclarationResolving
{
    public interface IUserComProjectSynchronizer
    {
        bool LastSyncOfUserComProjectsLoadedDeclarations { get; }
        IReadOnlyCollection<string> UserProjectIdsUnloaded { get; }

        void SyncUserComProjects();
    }
}