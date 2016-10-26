using System.IO;
using System.Reflection;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

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

        public string Export(IVBComponent component)
        {
            return component.ExportAsSourceFile(ExportPath);
        }
    }
}
