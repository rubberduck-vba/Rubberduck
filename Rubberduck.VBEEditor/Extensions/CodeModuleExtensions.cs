using System;
using System.Text.RegularExpressions;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.Extensions
{
    public static class CodeModuleExtensions
    {
        public static bool ContainsCompilationDirectives(this ICodeModule codeModule, Selection selection)
        {
            // We need to inspect the text directly from the code module because the parse tree do not have compilation
            // constants or directives easily available (they are in a hidden channel and are encapsulated). Fortunately,
            // finding compilation constants or directives are easy; they must be prefixed by a "#" and can only have
            // whitespace between it and the start of the line. Labels or line numbers are not legal in those lines.
            var rawCode = string.Join(Environment.NewLine,
                codeModule.GetLines(selection));
            var regex = new Regex(@"^(\s?)+#", RegexOptions.Multiline);
            return regex.Matches(rawCode).Count > 0;
        }

        public static bool ContainsCompilationDirectives(this ICodeModule codeModule)
        {
            var rawCode = string.Join(Environment.NewLine,
                codeModule.GetLines(1, codeModule.CountOfLines));
            var regex = new Regex(@"^(\s?)+#", RegexOptions.Multiline);
            return (regex.Matches(rawCode).Count > 0);
        }
    }
}
