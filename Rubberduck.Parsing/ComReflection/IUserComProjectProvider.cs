using System.Collections.Generic;

namespace Rubberduck.Parsing.ComReflection
{
    public interface IUserComProjectProvider
    {
        ComProject UserProject(string projectId);
        IReadOnlyDictionary<string, ComProject> UserProjects();
    }
}