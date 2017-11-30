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

        /// <summary>
        /// Determines whether two selections overlaps each other. This is useful for validating whether a required selection is wholly contained within a given selection.
        /// </summary>
        /// <param name="thisSelection">Target selection, usually representing user's selection</param>
        /// <param name="selection">The selection that might overlaps the target selection</param>
        /// <returns>Boolean with true indicating that the selections overlaps</returns>
        public static bool Overlaps(this Selection thisSelection, Selection selection)
        {
            if (thisSelection.StartLine == selection.EndLine)
            {
                if (thisSelection.StartColumn <= selection.EndColumn)
                    return true;
            }
            else if(thisSelection.EndLine == selection.EndLine)
            {
                if (thisSelection.EndColumn <= selection.EndColumn)
                    return true;
            }

            if (thisSelection.EndLine == selection.StartLine)
            {
                if (thisSelection.EndColumn >= selection.StartColumn)
                    return true;
            }
            else if(thisSelection.StartLine == selection.StartLine)
            {
                if (thisSelection.StartColumn <= selection.StartColumn)
                    return true;
            }

            if (thisSelection.StartLine < selection.EndLine && thisSelection.EndLine > selection.StartLine)
            {
                return true;
            }

            if (thisSelection.EndLine > selection.StartLine && thisSelection.StartLine < selection.EndLine)
            {
                return true;
            }

            return false; 
        }
    }
}
