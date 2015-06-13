using System.IO;
using NetOffice.VBIDEApi;
using NetOffice.VBIDEApi.Enums;

namespace Rubberduck.VBEditor.Extensions
{
    public static class VBComponentExtensions
    {
        internal const string ClassExtesnion = ".cls";
        internal const string FormExtension = ".frm";
        internal const string StandardExtension = ".bas";
        internal const string FormBinaryExtension = ".frx";
        internal const string DocClassExtension = ".doccls";

        /// <summary>
        /// Exports the component to the directoryPath. The file is name matches the component name and file extension is based on the component's type.
        /// </summary>
        /// <param name="component">The component to be exported to the file system.</param>
        /// <param name="directoryPath">Destination Path for the resulting source file.</param>
        public static void ExportAsSourceFile(this VBComponent component, string directoryPath)
        {
            string filePath = Path.Combine(directoryPath, component.Name + component.Type.FileExtension());
            if (component.Type == vbext_ComponentType.vbext_ct_Document)
            {
                int lineCount = component.CodeModule.CountOfLines;
                if (lineCount > 0)
                {
                    var text = component.CodeModule.get_Lines(1, lineCount);
                    File.WriteAllText(filePath, text);
                }
            }
            else
            {
                component.Export(filePath);
            }
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
                    return ClassExtesnion;
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