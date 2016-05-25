using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Parsing.Symbols
{
    public sealed class ProjectDeclaration : Declaration
    {
        private readonly List<ProjectReference> _projectReferences;

        public ProjectDeclaration(
            QualifiedMemberName qualifiedName,
            string name,
            bool isBuiltIn)
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
                  isBuiltIn)
        {
            _projectReferences = new List<ProjectReference>();
        }

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
    }
}
