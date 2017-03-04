using System.Runtime.InteropServices;
using System.Reflection;
using System.IO;
using System.Linq;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UnitTesting
{
    [ComVisible(false)]
    public static class ProjectTestExtensions
    {
        public static void EnsureReferenceToAddInLibrary(this IVBProject project)
        {
            var assembly = Assembly.GetExecutingAssembly();

            var name = assembly.GetName().Name.Replace('.', '_');
            var referencePath = Path.ChangeExtension(assembly.Location, ".tlb");

            var references = project.References;
            {
                var reference = references.SingleOrDefault(r => r.Name == name);
                if (reference != null)
                {
                    references.Remove(reference);
                }

                if (references.All(r => r.FullPath != referencePath))
                {
                    references.AddFromFile(referencePath);
                }
            }
        }
    }
}
