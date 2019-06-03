using System.Collections.Generic;

namespace Rubberduck.Parsing.ComReflection
{
    public interface IUserComProjectRepository : IUserComProjectProvider
    {
        void RefreshUserComProjects(IReadOnlyCollection<string> projectIdsToReload);
    }
}