using System.IO;
using System.Reflection;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Common
{
    public class ModuleExporter : IModuleExporter
    {
        public bool TempFile { get; private set; }

        public string ExportPath => TempFile
            ? ApplicationConstants.RUBBERDUCK_TEMP_PATH
            // note that App is not in the entry-point assembly, since Core is not the entry point anymore
            : Path.GetDirectoryName(Assembly.GetAssembly(typeof(App)).Location);

        public string Export(IVBComponent component, bool tempFile = true)
        {
            TempFile = tempFile;
            return component.ExportAsSourceFile(ExportPath, tempFile);
        }
    }
}
