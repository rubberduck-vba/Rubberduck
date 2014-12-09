using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;

namespace Rubberduck.Reflection
{
    [ComVisible(false)]
    public static class ProjectExtensions
    {
        public static IEnumerable<string> ComponentNames(this VBProject project)
        {
            foreach (VBComponent component in project.VBComponents)
            {
                yield return component.Name;
            }
        }

        public static void EnsureReferenceToAddInLibrary(this VBProject project)
        {
            var referencePath = System.IO.Path.ChangeExtension(System.Reflection.Assembly.GetExecutingAssembly().Location, ".tlb");
            //var existing = project.References.Cast<Reference>().Any(r => r.Name == "Rubberduck");

            List<Reference> existing = project.References.Cast<Reference>().Where(r => r.Name == "Rubberduck").ToList();
            foreach (Reference reference in existing)
            {
                project.References.Remove(reference);
            }

            if (project.References.Cast<Reference>().All(r => r.FullPath != referencePath))
            {
                project.References.AddFromFile(referencePath);
            }
        }
    }
}
