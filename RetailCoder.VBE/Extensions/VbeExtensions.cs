using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.VBEHost;

namespace Rubberduck.Extensions
{
    public static class VbeExtensions
    {
        public static CodeModule FindCodeModule(this VBE vbe, QualifiedModuleName qualifiedName)
        {
            return qualifiedName.Component.CodeModule;
        }

        public static void SetSelection(this VBE vbe, QualifiedSelection selection)
        {
            //not a very robust method. Breaks if there are multiple projects with the same name.
            var project = vbe.VBProjects.Cast<VBProject>()
                            .FirstOrDefault(p => p.Protection != vbext_ProjectProtection.vbext_pp_locked 
                                && p.Equals(selection.QualifiedName.Project));

            VBComponent component = null;
            if (project != null)
            {
                component = project.VBComponents.Cast<VBComponent>()
                                .FirstOrDefault(c => c.Equals(selection.QualifiedName.Component));
            }

            if (component == null)
            {
                return;
            }

            component.CodeModule.CodePane.SetSelection(selection.Selection);
        }

        public static CodeModuleSelection FindInstruction(this VBE vbe, QualifiedModuleName qualifiedModuleName, Selection selection)
        {
            var module = FindCodeModule(vbe, qualifiedModuleName);
            if (module == null)
            {
                return null;
            }

            return new CodeModuleSelection(module, selection);
        }

        /// <summary> Returns the type of Office Application that is hosting the VBE. </summary>
        /// <returns> Returns null if Unit Testing does not support Host Application.</returns>
        public static IHostApplication HostApplication(this VBE vbe)
        {
            foreach (var reference in vbe.ActiveVBProject.References.Cast<Reference>()
                .Where(reference => reference.BuiltIn && reference.Name != "VBA"))
            {
                switch (reference.Name)
                {
                    case "Excel":
                        return new ExcelApp();
                    case "Access":
                        return new AccessApp();
                    case "Word":
                        return new WordApp();
                    case "PowerPoint":
                        return new PowerPointApp();
                    case "Outlook":
                        return new OutlookApp();
                    case "Publisher":
                        return new PublisherApp();
                }
            }

            return null;
        }
    }
}