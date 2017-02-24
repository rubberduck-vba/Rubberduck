using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Common
{
    public static class CodeModuleExtensions
    {
        public static void ReplaceToken(this ICodeModule module, IToken token, string replacement)
        {
            var original = module.GetLines(token.Line, 1);
            var result = ReplaceStringAtIndex(original, token.Text, replacement, token.Column);
            module.ReplaceLine(token.Line, result);
        }

        public static void ReplaceIdentifierReferenceName(this ICodeModule module, IdentifierReference identifierReference, string replacement)
        {
            var original = module.GetLines(identifierReference.Selection.StartLine, 1);
            var result = ReplaceStringAtIndex(original, identifierReference.IdentifierName, replacement, identifierReference.Context.Start.Column);
            module.ReplaceLine(identifierReference.Selection.StartLine, result);
        }

        public static void InsertLines(this ICodeModule module, int startLine, string[] lines)
        {
            int lineNumber = startLine;
            for ( int idx = 0; idx < lines.Length; idx++ )
            {
                module.InsertLines(lineNumber, lines[idx]);
                lineNumber++;
            }
        }
        private static string ReplaceStringAtIndex(string original, string toReplace, string replacement, int startIndex)
        {
            var stopIndex = startIndex + toReplace.Length - 1;
            var prefix = original.Substring(0, startIndex);
            var suffix = (stopIndex >= original.Length) ? string.Empty : original.Substring(stopIndex + 1);

            if(original.Substring(startIndex, stopIndex - startIndex + 1).IndexOf(toReplace) != 0)
            {
                return original;
            }

            return prefix + toReplace.Replace(toReplace, replacement) + suffix;
        }
    }
}
