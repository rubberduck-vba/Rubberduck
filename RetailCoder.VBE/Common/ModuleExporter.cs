using System.IO;
using System.Reflection;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.VBA;

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
