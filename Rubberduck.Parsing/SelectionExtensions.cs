using Antlr4.Runtime;
using Rubberduck.VBEditor;
using System;

namespace Rubberduck.Parsing
{
    /// <summary>
    /// Provide extensions on selections & contexts/tokens
    /// to assist in validating whether a selection contains
    /// a given context or a token. 
    /// </summary>
    public static class SelectionExtensions
    {
        /// <summary>
        /// Validates whether a token is contained within a given Selection
        /// </summary>
        /// <param name="selection">One-based selection, usually from CodePane.Selection</param>
        /// <param name="token">An individual token within a module's parse tree</param>
        /// <returns>Boolean with true indicating that token is within the selection</returns>
        public static bool Contains(this Selection selection, IToken token)
        {
            return
                (((selection.StartLine == token.Line) && (selection.StartColumn - 1) <= token.Column) 
                    || (selection.StartLine < token.Line))
             && (((selection.EndLine == token.EndLine()) && (selection.EndColumn - 1) >= (token.EndColumn())) 
                    || (selection.EndLine > token.EndLine()));
        }

        /// <summary>
        /// Validates whether a context is contained within a given Selection
        /// </summary>
        /// <param name="context">A context which contains several tokens within a module's parse tree</param>
        /// <param name="selection">One-based selection, usually from CodePane.Selection</param>
        /// <returns>Boolean with true indicating that context is within the selection</returns>
        public static bool Contains(this ParserRuleContext context, Selection selection)
        {
            return
               (((selection.StartLine == context.Start.Line) && (selection.StartColumn - 1) <= context.Start.Column) 
                    || (selection.StartLine < context.Start.Line))
            && (((selection.EndLine == context.Stop.EndLine()) && (selection.EndColumn - 1) >= (context.Stop.EndColumn())) 
                    || (selection.EndLine > context.Stop.EndLine()));
        }

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
            if (token.Text.Contains(Environment.NewLine))
            {
                var splitStrings = token.Text.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
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
            if(token.Text.Contains(Environment.NewLine))
            {
                var splitStrings = token.Text.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);

                return token.Line + (splitStrings.Length - 1);
            }

            return token.Line;
        }
    }
}
