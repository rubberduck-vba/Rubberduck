using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Vbe.Interop;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.Parsing.VBA;

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
