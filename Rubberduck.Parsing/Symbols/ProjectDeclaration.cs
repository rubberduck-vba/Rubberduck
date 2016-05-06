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
            string name)
            : base(
                  qualifiedName,
                  null,
                  (Declaration)null, 
                  name,
                  false, 
                  false, 
                  Accessibility.Implicit, 
                  DeclarationType.Project,
                  null,
                  Selection.Home,
                  false)
        {
            _projectReferences = new List<ProjectReference>();
        }

        public IEnumerable<ProjectReference> ProjectReferences
        {
            get
            {
                return _projectReferences.OrderBy(reference => reference.Priority);
            }
        }

        public void AddProjectReference(string referencedProjectId, int priority)
        {
            _projectReferences.Add(new ProjectReference(referencedProjectId, priority));
        }
    }
}
