using System;
using System.Runtime.InteropServices;
using System.Reflection;
using Microsoft.Win32;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UnitTesting
{
    [ComVisible(false)]
    public static class ProjectTestExtensions
    {
        public static void EnsureReferenceToAddInLibrary(this IVBProject project)
        {
            var libFolder = IntPtr.Size == 8 ? "win64" : "win32";
            // TODO: This assumes the current assembly is same major/minor as the TLB!!!
            var libVersion = Assembly.GetExecutingAssembly().GetName().Version;
            const string libGuid = RubberduckGuid.RubberduckTypeLibGuid;
            var pathKey = Registry.ClassesRoot.OpenSubKey($@"TypeLib\{{{libGuid}}}\{libVersion.Major}.{libVersion.Minor}\0\{libFolder}");
            
            var referencePath = pathKey?.GetValue(string.Empty, string.Empty) as string;
            string name = null;

            if (!string.IsNullOrWhiteSpace(referencePath))
            {
                var tlbKey =
                    Registry.ClassesRoot.OpenSubKey($@"TypeLib\{{{libGuid}}}\{libVersion.Major}.{libVersion.Minor}");

                name = tlbKey?.GetValue(string.Empty, string.Empty) as string;
            }
            
            if (string.IsNullOrWhiteSpace(referencePath) || string.IsNullOrWhiteSpace(name))
            {
                throw new InvalidOperationException("Cannot locate the tlb in the registry or the entry may be corrupted. Therefore early binding is not possible");
            }

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
