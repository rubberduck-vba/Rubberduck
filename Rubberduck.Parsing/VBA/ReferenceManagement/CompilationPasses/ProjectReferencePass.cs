using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA.ReferenceManagement.CompilationPasses
{
    public sealed class ProjectReferencePass : ICompilationPass
    {
        private readonly DeclarationFinder _declarationFinder;

        public ProjectReferencePass(DeclarationFinder declarationFinder)
        {
            _declarationFinder = declarationFinder;
        }

        public void Execute(IReadOnlyCollection<QualifiedModuleName> modules)
        {
            var projects = _declarationFinder.Projects.Cast<ProjectDeclaration>().ToList();
            var allReferences = projects.Where(p => p.IsUserDefined).SelectMany(p => p.ProjectReferences).ToList();
            var builtInProjects = projects.Where(p => !p.IsUserDefined).ToList();
            // Give each built-in project access to all other projects so that e.g. CurrentDb in Access has access to the Database class defined in a different project.
            foreach (var builtInProject in builtInProjects)
            {
                builtInProject.ClearProjectReferences();    //Built-in projects have no project references of their own.
                foreach (var reference in allReferences)
                {
                    builtInProject.AddProjectReference(reference.ReferencedProjectId, reference.Priority);
                }
            }
        }
    }
}
