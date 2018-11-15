using System.Collections.Generic;

namespace Rubberduck.Parsing.ComReflection
{
    public interface IUserComProjectProvider
    {
        ComProject UserProject(string projectId);
        IEnumerable<(string projectId, ComProject comProject)> UserProjects();
    }
}