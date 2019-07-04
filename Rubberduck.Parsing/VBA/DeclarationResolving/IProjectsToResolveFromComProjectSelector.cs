using System.Collections.Generic;

namespace Rubberduck.Parsing.VBA.DeclarationResolving
{
    public interface IProjectsToResolveFromComProjectSelector
    {
        IReadOnlyCollection<string> ProjectsToResolveFromComProjects { get; }
        void RefreshProjectsToResolveFromComProjectSelector();
        bool ToBeResolvedFromComProject(string projectId);
    }
}