using System.IO;
using System.Reflection;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.DisposableWrappers;
using Rubberduck.VBEditor.DisposableWrappers.VBA;

namespace Rubberduck.Common
{
    public class ModuleExporter : IModuleExporter
    {
        public string ExportPath
        {
            get
            {
                var assemblyLocation = Assembly.GetAssembly(typeof(App)).Location;
                return Path.GetDirectoryName(assemblyLocation);
            }
        }

        public string Export(VBComponent component)
        {
            return component.ExportAsSourceFile(ExportPath);
        }
    }
}
