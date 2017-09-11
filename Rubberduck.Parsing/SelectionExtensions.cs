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
        public static bool Contains(this Selection selection, IToken token)
        {
            return
                (((selection.StartLine == token.Line) && (selection.StartColumn - 1) <= token.Column) || (selection.StartLine < token.Line))
             && (((selection.EndLine == token.Line) && (selection.EndColumn - 1) >= (token.EndColumn())) || (selection.EndLine > token.Line));
        }

        public static bool Contains(this ParserRuleContext context, Selection selection)
        {
            return
               (((selection.StartLine == context.Start.Line) && (selection.StartColumn - 1) <= context.Start.Column) || (selection.StartLine < context.Start.Line))
            && (((selection.EndLine == context.Stop.Line) && (selection.EndColumn - 1) >= (context.Stop.EndColumn())) || (selection.EndLine > context.Stop.Line));
        }

        /// <summary>
        /// Because a token can be spread across multiple lines with line continuations
        /// it is necessary to do some work to determine the token's actual ending column.
        /// Whitespace and newline should be preserved within the token.
        /// </summary>
        /// <param name="token">The last token within a given context to test</param>
        /// <returns></returns>
        public static int EndColumn(this IToken token)
        {
            if (token.Text.Contains(Environment.NewLine))
            {
                var splitStrings = token.Text.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
                var tokenOnLastLine = splitStrings[splitStrings.Length - 1];

                return tokenOnLastLine.Length;
            }
            else
            {
                return token.Column + token.Text.Length;
            }

        }
    }
}
