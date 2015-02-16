using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Vbe.Interop;
using Rubberduck.Inspections;
using Rubberduck.VBA;

namespace Rubberduck.Extensions
{
    public static class VBComponentsExtensions
    {
        public static QualifiedModuleName QualifiedName(this VBComponent component)
        {
            var moduleName = component.Name;
            var project = component.Collection.Parent;
            var hash = project.GetHashCode();
            var code = component.CodeModule.Lines().GetHashCode();

            return new QualifiedModuleName(project.Name, moduleName, hash, code);
        }

        /// <summary>
        /// Safely removes the specified VbComponent from the collection.
        /// </summary>
        /// <remarks>
        /// UserForms, Class modules, and Standard modules are completely removed from the project.
        /// Since Document type components can't be removed through the VBE, all code in its CodeModule are deleted instead.
        /// </remarks>
        public static void RemoveSafely(this VBComponents components, VBComponent component)
        {
            switch (component.Type)
            {
                case vbext_ComponentType.vbext_ct_ClassModule:
                case vbext_ComponentType.vbext_ct_StdModule:
                case vbext_ComponentType.vbext_ct_MSForm:
                    components.Remove(component);
                    break;
                case vbext_ComponentType.vbext_ct_ActiveXDesigner:
                case vbext_ComponentType.vbext_ct_Document:
                    component.CodeModule.Clear();
                    break;
                default:
                    break;
            }
        }

        public static void ImportSourceFile(this VBComponents components, string filePath)
        {
            var ext = System.IO.Path.GetExtension(filePath);
            var fileName = System.IO.Path.GetFileNameWithoutExtension(filePath);

            if (ext == VBComponentExtensions.DocClassExtension)
            {
                var component = components.Item(fileName);
                if (component != null)
                {
                    component.CodeModule.Clear();

                    var text = System.IO.File.ReadAllText(filePath);
                    component.CodeModule.AddFromString(text);
                }

            }
            else if(ext != VBComponentExtensions.FormBinaryExtension)
            {
                components.Import(filePath);
            }
        }
    }
}
