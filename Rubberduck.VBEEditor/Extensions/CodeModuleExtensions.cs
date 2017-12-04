using System;
using System.Text.RegularExpressions;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.Extensions
{
    public static class CodeModuleExtensions
    {
        // We need to inspect the text directly from the code module because the parse tree do not have compilation
        // constants or directives easily available (they are in a hidden channel and are encapsulated). Fortunately,
        // finding compilation constants or directives are easy; they must be prefixed by a "#" and can only have
        // whitespace between it and the start of the line. Labels or line numbers are not legal in those lines.

        /// <summary>
        /// Indicates whether a given selection within a code module contains conditional compilation directives which
        /// may impede refactoring of the code
        /// </summary>
        /// <param name="codeModule">CodeModule containing the selected code</param>
        /// <param name="selection">The selected code to test</param>
        /// <returns>True if any conditional compilation directives are found</returns>
        public static bool ContainsCompilationDirectives(this ICodeModule codeModule, Selection selection)
        {
            var rawCode = string.Join(Environment.NewLine,
                codeModule.GetLines(selection));
            return ContainsCompilationDirectives(rawCode);
        }

        /// <summary>
        /// Indicates whether the entire body of the code module contains conditional compilation directives which
        /// may impede refactoring of the code
        /// </summary>
        /// <param name="codeModule">CodeModule to test</param>
        /// <returns>True if any conditional compilation directives are found</returns>
        public static bool ContainsCompilationDirectives(this ICodeModule codeModule)
        {
            var rawCode = string.Join(Environment.NewLine,
                codeModule.GetLines(1, codeModule.CountOfLines));
            return ContainsCompilationDirectives(rawCode);
        }

        private static bool ContainsCompilationDirectives(string rawCode)
        {
            const string regexExpression = @"^\s*#";
            var regex = new Regex(regexExpression, RegexOptions.Multiline);
            return (regex.Matches(rawCode).Count > 0);
        }
    }
}
