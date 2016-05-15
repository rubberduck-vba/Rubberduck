using System.Diagnostics;
using System.Linq;

namespace Rubberduck.Parsing.Symbols
{
    public sealed class ProjectReferencePass : ICompilationPass
    {
        private readonly DeclarationFinder _declarationFinder;

        public ProjectReferencePass(DeclarationFinder declarationFinder)
        {
            _declarationFinder = declarationFinder;
        }

        public void Execute()
        {
            Stopwatch stopwatch = Stopwatch.StartNew();
            var projects = _declarationFinder.FindProjects();
            var allReferences = projects.Where(p => !p.IsBuiltIn).SelectMany(p => ((ProjectDeclaration)p).ProjectReferences).ToList();
            var builtInProjects = projects.Where(p => p.IsBuiltIn).ToList();
            // Give each built-in project access to all other projects so that e.g. CurrentDb in Access has access to the Database class defined in a different project.
            foreach (var builtInProject in builtInProjects)
            {
                foreach (var reference in allReferences)
                {
                    ((ProjectDeclaration)builtInProject).AddProjectReference(reference.ReferencedProjectId, reference.Priority);
                }
            }
            stopwatch.Stop();
            Debug.WriteLine("Built-in project references linked up in {0}ms.", stopwatch.ElapsedMilliseconds);
        }
    }
}