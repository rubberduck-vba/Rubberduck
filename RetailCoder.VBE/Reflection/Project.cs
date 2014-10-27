using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Vbe.Interop;
using IDE = Microsoft.Vbe.Interop.VBE;
using System.Windows.Forms;

namespace RetailCoderVBE.Reflection
{
    internal static class ProjectExtensions
    {
        public static IEnumerable<string> ComponentNames(this VBProject project)
        {
            foreach (VBComponent component in project.VBComponents)
            {
                yield return component.Name;
            }
        }

        public static void EnsureReferenceToRetailCoderVBE(this VBProject project)
        {
            var referencePath = System.IO.Path.ChangeExtension(System.Reflection.Assembly.GetExecutingAssembly().Location, ".tlb");
            if (!project.References.Cast<Reference>().Any(r => r.FullPath == referencePath))
            {
                project.References.AddFromFile(referencePath);
            }
        }
    }
}
