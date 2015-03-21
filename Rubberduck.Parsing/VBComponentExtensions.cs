using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Vbe.Interop;

namespace Rubberduck.Parsing
{
    public static class VBComponentExtensions
    {
        public static QualifiedModuleName QualifiedName(this VBComponent component)
        {
            var moduleName = component.Name;
            var project = component.Collection.Parent;
            var hash = project.GetHashCode();
            var code = component.CodeModule.Lines().GetHashCode();

            return new QualifiedModuleName(project.Name, moduleName, hash, code);
        }

        public static string Lines(this CodeModule module)
        {
            if (module.CountOfLines == 0)
            {
                return string.Empty;
            }

            return module.Lines[1, module.CountOfLines];
        }
    }
}
