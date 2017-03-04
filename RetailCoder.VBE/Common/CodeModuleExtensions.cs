using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Common
{
    public static class CodeModuleExtensions
    {
        /// <summary>
        /// Removes a <see cref="Declaration"/> and its <see cref="Declaration.References"/>.
        /// </summary>
        /// <param name="module">The <see cref="ICodeModule"/> to modify.</param>
        /// <param name="target"></param>
        public static void Remove(this ICodeModule module, Declaration target)
        {
            var multipleDeclarations = target.DeclarationType == DeclarationType.Variable && target.HasMultipleDeclarationsInStatement();
            var context = GetStmtContext(target);
            var declarationText = context.GetText().Replace(" _" + Environment.NewLine, string.Empty);
            var selection = GetStmtContextSelection(target);

            var oldLines = module.GetLines(selection);

            var newLines = oldLines.Replace(" _" + Environment.NewLine, string.Empty)
                .Remove(selection.StartColumn - 1, declarationText.Length);

            if (multipleDeclarations)
            {
                selection = GetStmtContextSelection(target);
                newLines = RemoveExtraComma(module.GetLines(selection).Replace(oldLines, newLines),
                    target.CountOfDeclarationsInStatement(), target.IndexOfVariableDeclarationInStatement());
            }

            var newLinesWithoutExcessSpaces = newLines.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            for (var i = 0; i < newLinesWithoutExcessSpaces.Length; i++)
            {
                newLinesWithoutExcessSpaces[i] = newLinesWithoutExcessSpaces[i].RemoveExtraSpacesLeavingIndentation();
            }

            for (var i = newLinesWithoutExcessSpaces.Length - 1; i >= 0; i--)
            {
                if (newLinesWithoutExcessSpaces[i].Trim() == string.Empty)
                {
                    continue;
                }

                if (newLinesWithoutExcessSpaces[i].EndsWith(" _"))
                {
                    newLinesWithoutExcessSpaces[i] =
                        newLinesWithoutExcessSpaces[i].Remove(newLinesWithoutExcessSpaces[i].Length - 2);
                }
                break;
            }

            // remove all lines with only whitespace
            newLinesWithoutExcessSpaces = newLinesWithoutExcessSpaces.Where(str => str.Any(c => !char.IsWhiteSpace(c))).ToArray();

            module.DeleteLines(selection);
            if (newLinesWithoutExcessSpaces.Any())
            {
                module.InsertLines(selection.StartLine, string.Join(Environment.NewLine, newLinesWithoutExcessSpaces));
            }
        }

        private static Selection GetStmtContextSelection(Declaration target)
        {
            switch (target.DeclarationType)
            {
                case DeclarationType.Variable:
                    return target.GetVariableStmtContextSelection();
                case DeclarationType.Constant:
                    return target.GetConstStmtContextSelection();
                default:
                    return target.Context.GetSelection();
            }
        }

        private static ParserRuleContext GetStmtContext(Declaration target)
        {
            switch (target.DeclarationType)
            {
                case DeclarationType.Variable:
                    return target.GetVariableStmtContext();
                case DeclarationType.Constant:
                    return target.GetConstStmtContext();
                default:
                    return target.Context;
            }
        }

        private static string RemoveExtraComma(string str, int numParams, int indexRemoved)
        {
            #region usage example
            // Example use cases for this method (fields and variables):
            // Dim fizz as Boolean, dizz as Double
            // Private fizz as Boolean, dizz as Double
            // Public fizz as Boolean, _
            //        dizz as Double
            // Private fizz as Boolean _
            //         , dizz as Double _
            //         , iizz as Integer

            // Before this method is called, the parameter to be removed has 
            // already been removed.  This means 'str' will look like:
            // Dim fizz as Boolean, 
            // Private , dizz as Double
            // Public fizz as Boolean, _
            //        
            // Private  _
            //         , dizz as Double _
            //         , iizz as Integer

            // This method is responsible for removing the redundant comma
            // and returning a string similar to:
            // Dim fizz as Boolean
            // Private dizz as Double
            // Public fizz as Boolean _
            //        
            // Private  _
            //          dizz as Double _
            //         , iizz as Integer
            #endregion
            var commaToRemove = numParams == indexRemoved ? indexRemoved - 1 : indexRemoved;
            return str.Remove(str.NthIndexOf(',', commaToRemove), 1);
        }

        public static void Remove(this ICodeModule module, ConstantDeclaration target)
        {

        }

        public static void Remove(this ICodeModule module, IdentifierReference target)
        {
            var parent = target.Context.Parent as ParserRuleContext;
            module.Remove(parent.GetSelection(), parent);
        }

        public static void Remove(this ICodeModule module, IEnumerable<IdentifierReference> targets)
        {
            foreach (var target in targets.OrderByDescending(e => e.Selection))
            {
                module.Remove(target);
            }
        }

        public static void Remove(this ICodeModule module, Selection selection, ParserRuleContext instruction)
        {
            var originalCodeLines = module.GetLines(selection.StartLine, selection.LineCount);
            var originalInstruction = instruction.GetText();
            module.DeleteLines(selection.StartLine, selection.LineCount);

            var newCodeLines = originalCodeLines.Replace(originalInstruction, string.Empty);
            if (!string.IsNullOrEmpty(newCodeLines))
            {
                module.InsertLines(selection.StartLine, newCodeLines);
            }
        }

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
            var lineNumber = startLine;
            for (var idx = 0; idx < lines.Length; idx++)
            {
                module.InsertLines(lineNumber, lines[idx]);
                lineNumber++;
            }
        }

        private static string ReplaceStringAtIndex(string original, string toReplace, string replacement, int startIndex)
        {
            var modifiedContent = original.Remove(startIndex, toReplace.Length);
            return modifiedContent.Insert(startIndex, replacement);
        }
    }
}
