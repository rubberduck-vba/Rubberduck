using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Vbe.Interop;

namespace Rubberduck.Extensions
{
    public static class CodeModuleExtensions
    {
        /// <summary>
        /// Deletes all lines from the CodeModule
        /// </summary>
        public static void Clear(this CodeModule module)
        {
            module.DeleteLines(1, module.CountOfLines);
        }
    }
}
