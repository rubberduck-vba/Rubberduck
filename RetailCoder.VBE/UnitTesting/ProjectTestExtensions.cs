using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using System.Reflection;
using System.IO;

namespace Rubberduck.UnitTesting
{
    [ComVisible(false)]
    public static class ProjectTestExtensions
    {
        public static void EnsureReferenceToAddInLibrary(this VBProject project)
        {
            var assembly = Assembly.GetExecutingAssembly();

            var name = assembly.GetName().Name.Replace('.', '_');
            var referencePath = Path.ChangeExtension(assembly.Location, ".tlb");

            var references = project.References.Cast<Reference>().ToList();

            var reference = references.SingleOrDefault(r => r.Name == name);
            if (reference != null)
            {
                references.Remove(reference);
                project.References.Remove(reference);
            }

            if (references.All(r => r.FullPath != referencePath))
            {
                project.References.AddFromFile(referencePath);
            }
        }
    }
}
