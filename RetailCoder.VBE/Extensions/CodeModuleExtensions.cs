using System.Linq;
using Microsoft.Vbe.Interop;

namespace Rubberduck.Extensions
{
    /// <summary>
    /// VBE CodeModule extension methods.
    /// </summary>
    public static class CodeModuleExtensions
    {
        /// <summary>
        /// Gets an array of strings where each element is a line of code in the module.
        /// </summary>
        public static string[] Code(this CodeModule module)
        {
            var lines = module.CountOfLines;
            if (lines == 0)
            {
                return new string[]{};
            }

            return module.get_Lines(1, lines).Replace("\r", string.Empty).Split('\n');
        }
    }
}