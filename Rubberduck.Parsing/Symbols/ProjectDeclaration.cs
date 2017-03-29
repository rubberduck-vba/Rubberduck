using Rubberduck.Parsing.ComReflection;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.Symbols
{
    public sealed class ProjectDeclaration : Declaration
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
            MajorVersion = project.MajorVersion;
            MinorVersion = project.MinorVersion;
        }

        public long MajorVersion { get; set; }
        public long MinorVersion { get; set; }

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
        public override IVBProject Project => _project;

        public void AddProjectReference(string referencedProjectId, int priority)
        {
            if (_projectReferences.Any(p => p.ReferencedProjectId == referencedProjectId))
            {
                return;
            }
            _projectReferences.Add(new ProjectReference(referencedProjectId, priority));
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
                _displayName = _project != null ? _project.ProjectDisplayName : string.Empty;
                return _displayName;
            }
        }
    }
}
