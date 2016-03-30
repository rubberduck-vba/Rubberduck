using System;
using System.Collections.Generic;
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
            switch (component.Type)
            {
                case vbext_ComponentType.vbext_ct_MSForm:
                    ExportUserFormModule(component, path);
                    break;
                case vbext_ComponentType.vbext_ct_Document:
                    ExportDocumentModule(component, path);
                    break;
                default:
                    component.Export(path);
                    break;
            }

            return path;
        }

        private static void ExportUserFormModule(VBComponent component, string path)
        {
            // VBIDE API inserts an extra newline when exporting a UserForm module.
            // this issue causes forms to always be treated as "modified" in source control, which causes conflicts.
            // we need to remove the extra newline before the file gets written to its output location.

            var visibleCode = component.CodeModule.Lines().Split(new []{Environment.NewLine}, StringSplitOptions.None);
            var legitEmptyLineCount = visibleCode.TakeWhile(string.IsNullOrWhiteSpace).Count();

            var tempFile = component.ExportToTempFile();
            var contents = File.ReadAllLines(tempFile);
            var nonAttributeLines = contents.TakeWhile(line => !line.StartsWith("Attribute")).Count();
            var attributeLines = contents.Skip(nonAttributeLines).TakeWhile(line => line.StartsWith("Attribute")).Count();
            var declarationsStartLine = nonAttributeLines + attributeLines + 1;

            var emptyLineCount = contents.Skip(declarationsStartLine - 1)
                                         .TakeWhile(string.IsNullOrWhiteSpace)
                                         .Count();

            var code = contents;
            if (emptyLineCount > legitEmptyLineCount)
            {
                code = contents.Take(declarationsStartLine).Union(
                       contents.Skip(declarationsStartLine + emptyLineCount - legitEmptyLineCount))
                               .ToArray();
            }
            File.WriteAllLines(path, code);
        }

        private static void ExportDocumentModule(VBComponent component, string path)
        {
            var lineCount = component.CodeModule.CountOfLines;
            if (lineCount > 0)
            {
                var text = component.CodeModule.Lines[1, lineCount];
                File.WriteAllText(path, text);
            }
        }

        private static string ExportToTempFile(this VBComponent component)
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