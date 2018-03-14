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

            using (var references = project.References)
            {
                var reference = FindReferenceByName(references, name);
                if (reference != null)
                {
                    references.Remove(reference);
                    reference.Dispose();
                }

                if (!ReferenceWithPathExists(references, referencePath))
                {
                    references.AddFromFile(referencePath);
                }
            }
        }

        private static IReference FindReferenceByName(IReferences refernences, string name)
        {
            foreach(var reference in refernences)
            {
                if (reference.Name == name)
                {
                    return reference;
                }
                reference.Dispose();
            }

            return null;
        }

        private static bool ReferenceWithPathExists(IReferences refereneces, string path)
        {
            foreach (var reference in refereneces)
            {
                var referencePath = reference.FullPath;
                reference.Dispose();
                if (referencePath == path)
                {
                    return true;
                }
            }
            return false;
        }
    }
}
