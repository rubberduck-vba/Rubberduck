using System;
using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.ComManagement
{
    public interface IProjectsProvider : IDisposable
    {
        IVBProjects ProjectsCollection();
        IEnumerable<(string ProjectId, IVBProject Project)> Projects();
        IEnumerable<(string ProjectId, IVBProject Project)> LockedProjects();
        IVBProject Project(string projectId);
        IVBComponents ComponentsCollection(string projectId);
        IEnumerable<(QualifiedModuleName QualifiedModuleName, IVBComponent Component)> Components();
        IEnumerable<(QualifiedModuleName QualifiedModuleName, IVBComponent Component)> Components(string projectId);
        IVBComponent Component(QualifiedModuleName qualifiedModuleName);
    }
}
