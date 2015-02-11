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
                return new string[] { };
            }

            return module.get_Lines(1, lines).Replace("\r", string.Empty).Split('\n');
        }

        /// <summary>
        /// Returns all of the code in a module as a string.
        /// </summary>
        public static string Lines(this CodeModule module)
        {
            if (module.CountOfLines == 0)
            {
                return string.Empty;
            }

            return module.Lines[1, module.CountOfLines];
        }

        /// <summary>
        /// Deletes all lines from the CodeModule
        /// </summary>
        public static void Clear(this CodeModule module)
        {
            module.DeleteLines(1, module.CountOfLines);
        }
    }
}