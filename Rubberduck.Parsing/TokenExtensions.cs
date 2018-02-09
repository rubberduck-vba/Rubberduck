using System;
using Antlr4.Runtime;

namespace Rubberduck.Parsing
{
    public static class TokenExtensions
    {
        /// <summary>
        /// Obtain the actual last column the token occupies. Because a token can be spread 
        /// across multiple lines with line continuations it is necessary to do some work 
        /// to determine the token's actual ending column.
        /// Whitespace and newline should be preserved within the token.
        /// </summary>
        /// <param name="token">The last token within a given context to test</param>
        /// <returns>Zero-based column position</returns>
        public static int EndColumn(this IToken token)
        {
            if (token.Text == Environment.NewLine || token.Type == TokenConstants.Eof)
            {
                return token.Column;
            }
            else if (token.Text.Contains(Environment.NewLine))
            {
                var splitStrings = token.Text.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
                var lastOccupiedLine = splitStrings[splitStrings.Length - 1];

                return lastOccupiedLine.Length;
            }
            else
            {
                return token.Column + token.Text.Length;
            }
        }

        /// <summary>
        /// Obtain the actual last line token occupies. Typically it is same as token.Line but 
        /// when it contains line continuation and is spread across lines, extra newlines are
        /// counted and added.
        /// </summary>
        /// <param name="token"></param>
        /// <returns>One-based line position</returns>
        public static int EndLine(this IToken token)
        {
            if (token.Text != Environment.NewLine && token.Text.Contains(Environment.NewLine))
            {
                var splitStrings = token.Text.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);

                return token.Line + (splitStrings.Length - 1);
            }

            return token.Line;
        }
    }
}
