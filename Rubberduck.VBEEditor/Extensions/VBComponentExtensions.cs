using System;
using System.Linq;
using System.IO;
using Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.Extensions
{
    public static class VBComponentExtensions
    {
        internal const string ClassExtension = ".cls";
        internal const string FormExtension = ".frm";
        internal const string StandardExtension = ".bas";
        internal const string FormBinaryExtension = ".frx";
        internal const string DocClassExtension = ".doccls";

        /// <summary>
        /// Exports the component to the directoryPath. The file is name matches the component name and file extension is based on the component's type.
        /// </summary>
        /// <param name="component">The component to be exported to the file system.</param>
        /// <param name="directoryPath">Destination Path for the resulting source file.</param>
        public static string ExportAsSourceFile(this VBComponent component, string directoryPath)
        {
            var path = Path.Combine(directoryPath, component.Name + component.Type.FileExtension());
            if (component.Type == vbext_ComponentType.vbext_ct_Document)
            {
                var lineCount = component.CodeModule.CountOfLines;
                if (lineCount > 0)
                {
                    var text = component.CodeModule.get_Lines(1, lineCount);
                    File.WriteAllText(path, text);
                }
            }
            else if(component.Type == vbext_ComponentType.vbext_ct_MSForm)
            {
                // VBIDE API inserts an extra newline when exporting a UserForm module.
                // this issue causes forms to always be treated as "modified" in source control, which causes conflicts.
                // we need to remove the extra newline before the file gets written to its output location.
                var tempFile = component.ExportToTempFile();
                var contents = File.ReadAllLines(tempFile);
                var vbExposedAttributeLine = 0;
                for (var i = 0; i < contents.Length; i++)
                {
                    if (contents[i].StartsWith("Attribute VB_Exposed = "))
                    {
                        vbExposedAttributeLine = i;
                        break;
                    }
                }
                if (contents.Length > vbExposedAttributeLine && contents[vbExposedAttributeLine + 1].Replace(Environment.NewLine, string.Empty).Trim() == string.Empty)
                {
                    // first line in module is empty - assume it wasn't in the VBE
                    var edited = contents.Take(vbExposedAttributeLine).Union(contents.Skip(vbExposedAttributeLine + 1)).ToArray();
                    File.WriteAllLines(path, edited);
                }
            }
            else
            {
                // other module types can just be exported directly
                component.Export(path);
            }

            return path;
        }

        public static string ExportToTempFile(this VBComponent component)
        {
            var path = Path.Combine(Path.GetTempPath(), component.Name + component.Type.FileExtension());
            component.Export(path);
            return path;
        }

        /// <summary>
        /// Returns the proper file extension for the Component Type.
        /// </summary>
        /// <remarks>Document classes should properly have a ".cls" file extension.
        /// However, because they cannot be removed and imported like other component types, we need to make a distinction.</remarks>
        /// <param name="componentType"></param>
        /// <returns>File extension that includes a preceeding "dot" (.) </returns>
        public static string FileExtension(this vbext_ComponentType componentType)
        {
            switch (componentType)
            {
                case vbext_ComponentType.vbext_ct_ClassModule:
                    return ClassExtension;
                case vbext_ComponentType.vbext_ct_MSForm:
                    return FormExtension;
                case vbext_ComponentType.vbext_ct_StdModule:
                    return StandardExtension;
                case vbext_ComponentType.vbext_ct_Document:
                    // documents should technically be a ".cls", but we need to be able to tell them apart.
                    return DocClassExtension;
                case vbext_ComponentType.vbext_ct_ActiveXDesigner:
                default:
                    return string.Empty;
            }
        }
    }
}