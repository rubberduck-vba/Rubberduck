using Antlr4.Runtime;
using Rubberduck.VBEditor;
using System.Linq;

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
        /// Exposes LINQ method to semantic predicates.
        /// </summary>
        public static bool Contains(this int[] values, int value)
        {
            return Enumerable.Contains(values, value);
        }

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
        /// <param name="selection">One-based selection, usually from CodePane.Selection</param>
        /// <param name="context">A context which contains several tokens within a module's parse tree</param>
        /// <returns>Boolean with true indicating that context is within the selection</returns>
        public static bool Contains(this Selection selection, ParserRuleContext context)
        {
            return
               (((selection.StartLine == context.Start.Line) && (selection.StartColumn - 1) <= context.Start.Column) 
                    || (selection.StartLine < context.Start.Line))
            && (((selection.EndLine == context.Stop.EndLine()) && (selection.EndColumn - 1) >= (context.Stop.EndColumn())) 
                    || (selection.EndLine > context.Stop.EndLine()));
        }

        /// <summary>
        /// Convenience method for validating that a selection is inside a specified parser rule context.
        /// </summary>
        /// <param name="selection">The selection that should be contained within the ParserRuleContext</param>
        /// <param name="context">The containing ParserRuleContext</param>
        /// <returns>Boolean with true indicating that the selection is inside the given context</returns>
        public static bool IsContainedIn(this Selection selection, ParserRuleContext context)
        {
            return context.GetSelection().Contains(selection);
        }
    }
}
