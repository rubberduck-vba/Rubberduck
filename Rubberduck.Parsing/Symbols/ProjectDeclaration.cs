using System;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Parsing.Symbols
{
    public sealed class ProjectDeclaration : Declaration, IDisposable
    {
        private readonly List<ProjectReference> _projectReferences;

        public ProjectDeclaration(
            QualifiedMemberName qualifiedName,
            string name,
            bool isUserDefined)
            : base(
                  qualifiedName,
                  null,
                  (Declaration)null,
                  name,
                  null,
                  false,
                  false,
                  Accessibility.Implicit,
                  DeclarationType.Project,
                  null,
                  null,
                  Selection.Home,
                  false,
                  null,
                  isUserDefined)
        {
            _projectReferences = new List<ProjectReference>();
        }

        public ProjectDeclaration(ComProject project, QualifiedModuleName module)
            : this(module.QualifyMemberName(project.Name), project.Name, false)
        {
            Guid = project.Guid;
            MajorVersion = project.MajorVersion;
            MinorVersion = project.MinorVersion;
        }

        public Guid Guid { get; }
        public long MajorVersion { get; }
        public long MinorVersion { get; }

        public IReadOnlyList<ProjectReference> ProjectReferences
        {
            get
            {
                return _projectReferences.OrderBy(reference => reference.Priority).ToList();
            }
        }

        public void AddProjectReference(string referencedProjectId, int priority)
        {
            if (_projectReferences.Any(p => p.ReferencedProjectId == referencedProjectId))
            {
                return;
            }
            _projectReferences.Add(new ProjectReference(referencedProjectId, priority));
        }

        public void ClearProjectReferences()
        {
            _projectReferences.Clear();
        }

        public bool IsDisposed { get; private set; }

        public void Dispose()
        {
            IsDisposed = true;
        }
    }
}
