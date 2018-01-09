using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.ComManagement
{
    public interface IProjectsProvider
    {
        IVBProjects ProjectsCollection();
        IEnumerable<(string ProjectId, IVBProject Project)> Projects();
        IVBProject Project(string projectId);
        IVBComponents ComponentsCollection(string projectId);
        IEnumerable<(QualifiedModuleName QualifiedModuleName, IVBComponent Component)> Components(string projectId = null);
        IVBComponent Component(QualifiedModuleName qualifiedModuleName);
    }
}
