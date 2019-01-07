using System;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.Symbols
{
    public sealed class ProjectDeclaration : Declaration, IDisposable
    {
        private readonly List<ProjectReference> _projectReferences;

        public ProjectDeclaration(
            QualifiedMemberName qualifiedName,
            string name,
            bool isUserDefined,
            IVBProject project)
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
            _project = project;
            _projectReferences = new List<ProjectReference>();
        }

        public ProjectDeclaration(ComProject project, QualifiedModuleName module)
            : this(module.QualifyMemberName(project.Name), project.Name, false, null)
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

        private readonly IVBProject _project;
        /// <summary>
        /// Gets a reference to the VBProject the declaration is made in.
        /// </summary>
        /// <remarks>
        /// This property is intended to differenciate identically-named VBProjects.
        /// </remarks>
        public override IVBProject Project => IsDisposed ? null : _project;

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

        private string _displayName;
        /// <summary>
        /// WARNING: This property has side effects. It changes the ActiveVBProject, which causes a flicker in the VBE.
        /// This should only be called if it is *absolutely* necessary.
        /// </summary>
        public override string ProjectDisplayName
        {
            get
            {
                if (_displayName != null)
                {
                    return _displayName;
                }
                _displayName = !IsDisposed && _project != null ? _project.ProjectDisplayName : string.Empty;
                return _displayName;
            }
        }


        public bool IsDisposed { get; private set; }

        public void Dispose()
        {
            IsDisposed = true;
        }
    }
}
