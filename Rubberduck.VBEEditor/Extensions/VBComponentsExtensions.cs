using System;
using System.IO;
using System.Linq;
using Microsoft.Vbe.Interop;

// todo: untangle this mess

namespace Rubberduck.VBEditor.Extensions
{
    public static class VBComponentsExtensions
    {
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
            var ext = Path.GetExtension(filePath);
            var name = Path.GetFileNameWithoutExtension(filePath);
            if (!File.Exists(filePath))
            {
                return;
            }

            var codeString = File.ReadAllText(filePath);
            var codeLines = codeString.Split(new []{Environment.NewLine}, StringSplitOptions.None);
            if (ext == VBComponentExtensions.DocClassExtension)
            {
                var component = components.Item(name);
                if (component != null)
                {
                    component.CodeModule.Clear();
                    component.CodeModule.AddFromString(codeString);
                }
            }
            else if(ext != VBComponentExtensions.FormBinaryExtension)
            {
                components.Import(filePath);
            }

            if (ext == VBComponentExtensions.FormExtension)
            {
                var component = components.Item(name);
                // note: vbeCode contains an extraneous line here:
                //var vbeCode = component.CodeModule.Lines().Split(new []{Environment.NewLine}, StringSplitOptions.None);
                
                var nonAttributeLines = codeLines.TakeWhile(line => !line.StartsWith("Attribute")).Count();
                var attributeLines = codeLines.Skip(nonAttributeLines).TakeWhile(line => line.StartsWith("Attribute")).Count();
                var declarationsStartLine = nonAttributeLines + attributeLines + 1;
                var correctCodeString = string.Join(Environment.NewLine, codeLines.Skip(declarationsStartLine - 1).ToArray());

                component.CodeModule.Clear();
                component.CodeModule.AddFromString(correctCodeString);
            }
        }
    }
}
