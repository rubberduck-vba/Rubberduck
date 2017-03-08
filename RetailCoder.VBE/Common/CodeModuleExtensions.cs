using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
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
            if (!module.Equals(target.QualifiedName.QualifiedModuleName.Component.CodeModule))
            {
                throw new ArgumentException("Target is not declared in specified module.");
            }

            var sortedItems = target.References
                .Where(reference => module.Equals(reference.QualifiedModuleName.Component.CodeModule))
                .Select(reference => Tuple.Create((object) reference, reference.Selection))
                .Concat(new[] {Tuple.Create((object) target, target.Selection)})
                .OrderByDescending(t => t.Item2);

            foreach (var tuple in sortedItems)
            {
                if (tuple.Item1 is Declaration)
                {
                    RemoveDeclarationOnly(module, target);
                }
                else
                {
                    var reference = (IdentifierReference) tuple.Item1;
                    Remove(reference.QualifiedModuleName.Component.CodeModule, reference);
                }
            }
        }

        private static void RemoveDeclarationOnly(this ICodeModule module, Declaration target)
        {
            var multipleDeclarations = target.DeclarationType == DeclarationType.Variable && target.HasMultipleDeclarationsInStatement();
            var context = GetStmtContext(target);
            var declarationText = context.GetText().Replace(" _" + Environment.NewLine, Environment.NewLine);
            var selection = GetStmtContextSelection(target);
            Debug.Assert(selection.StartColumn > 0);

            var oldLines = module.GetLines(selection);
            var indent = oldLines.IndexOf(oldLines.FirstOrDefault(c => c != ' ')) + 1;

            var newLines = oldLines
                .Replace(" _" + Environment.NewLine, Environment.NewLine)
                .Remove(selection.StartColumn - 1, declarationText.Length - selection.StartColumn + indent);

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

        public static void Remove(this ICodeModule module, IdentifierReference target)
        {
            var parent = (ParserRuleContext)target.Context.Parent;
            if (target.IsAssignment)
            {
                // target is LHS of assignment; need to know if there's a procedure call in RHS
                var letStmt = parent as VBAParser.LetStmtContext;
                var setStmt = parent as VBAParser.SetStmtContext;

                string argList;
                if (HasProcedureCall(letStmt, out argList) || HasProcedureCall(setStmt, out argList))
                {
                    // need to remove LHS only; RHS expression may have side-effects
                    var original = parent.GetText();
                    var replacement = ReplaceStringAtIndex(original, target.IdentifierName + " = ", string.Empty, 0);
                    if (argList != null)
                    {
                        var atIndex = replacement.IndexOf(argList, StringComparison.OrdinalIgnoreCase);
                        var plainArgs = " " + argList.Substring(1, argList.Length - 2);
                        replacement = ReplaceStringAtIndex(replacement, argList, plainArgs, atIndex);
                    }
                    module.ReplaceLine(parent.Start.Line, replacement);
                    return;
                }
            }

            module.Remove(parent.GetSelection(), parent);
        }

        private static bool HasProcedureCall(VBAParser.LetStmtContext context, out string argList)
        {
            if (context == null)
            {
                argList = null;
                return false;
            }
            return HasProcedureCall(context.expression(), out argList);
        }

        private static bool HasProcedureCall(VBAParser.SetStmtContext context, out string argList)
        {
            if (context == null)
            {
                argList = null;
                return false;
            }
            return HasProcedureCall(context.expression(), out argList);
        }

        private static bool HasProcedureCall(VBAParser.ExpressionContext context, out string argList)
        {
            // bug: what if complex expression has multiple arg lists?
            argList = GetArgListString(context.FindChildren<VBAParser.ArgListContext>().FirstOrDefault())
                      ?? GetArgListString(context.FindChildren<VBAParser.ArgumentListContext>().FirstOrDefault());

            return !(context is VBAParser.LiteralExprContext 
                  || context is VBAParser.NewExprContext
                  || context is VBAParser.BuiltInTypeExprContext);
        }

        private static string GetArgListString(VBAParser.ArgListContext context)
        {
            return context == null ? null : context.GetText();
        }

        private static string GetArgListString(VBAParser.ArgumentListContext context)
        {
            return context == null ? null : "(" + context.GetText() + ")";
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
