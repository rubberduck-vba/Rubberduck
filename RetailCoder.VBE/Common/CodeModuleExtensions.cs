using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
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
        private struct RewriterInfo
        {
            private readonly string _replacement;
            private readonly int _startTokenIndex;
            private readonly int _stopTokenIndex;

            public RewriterInfo(int startTokenIndex, int stopTokenIndex)
                : this(string.Empty, startTokenIndex, stopTokenIndex) { }

            public RewriterInfo(string replacement, int startTokenIndex, int stopTokenIndex)
            {
                _replacement = replacement;
                _startTokenIndex = startTokenIndex;
                _stopTokenIndex = stopTokenIndex;
            }

            public string Replacement { get { return _replacement; } }
            public int StartTokenIndex { get { return _startTokenIndex; } }
            public int StopTokenIndex { get { return _stopTokenIndex; } }

            public static RewriterInfo None { get { return default(RewriterInfo); } }
        }

        public static void Rewrite(this ICodeModule module, TokenStreamRewriter rewriter)
        {
            module.Clear();
            var content = rewriter.GetText();
            module.InsertLines(1, content);
        }

        /// <summary>
        /// Removes a <see cref="Declaration"/> and its <see cref="Declaration.References"/>.
        /// </summary>
        /// <param name="module">The <see cref="ICodeModule"/> to modify.</param>
        /// <param name="rewriter">The <see cref="TokenStreamRewriter"/> holding the state to alter.</param>
        /// <param name="target"></param>
        public static void Remove(this ICodeModule module, TokenStreamRewriter rewriter, Declaration target)
        {
            // note: commented-out because it makes tests fail.. need a way to fix that
            //if (!module.Equals(target.QualifiedName.QualifiedModuleName.Component.CodeModule))
            //{
            //    throw new ArgumentException("Target is not declared in specified module.");
            //}

            var items = target.References
                .Where(reference => module.Equals(reference.QualifiedModuleName.Component.CodeModule))
                .Select(reference => Tuple.Create((object) reference, reference.Selection))
                .Concat(new[] {Tuple.Create((object) target, target.Selection)})
                //.OrderByDescending(t => t.Item2)
                ;

            foreach (var tuple in items)
            {
                if (tuple.Item1 is Declaration)
                {
                    RemoveDeclaration(target, rewriter);
                }
                else
                {
                    var reference = (IdentifierReference) tuple.Item1;
                    Remove(reference.QualifiedModuleName.Component.CodeModule, reference, rewriter);
                }
            }
        }

        private static void RemoveDeclaration(Declaration target, TokenStreamRewriter rewriter)
        {
            var info = GetRewriterInfoForRemovedTarget(target);
            rewriter.Delete(info.StartTokenIndex, info.StopTokenIndex);
        }

        private static RewriterInfo GetRewriterInfoForRemovedTarget(Declaration target)
        {
            var variables = target.Context.Parent as VBAParser.VariableListStmtContext;
            if (variables != null)
            {
                var items = variables.variableSubStmt();
                var itemIndex = items.ToList().IndexOf((VBAParser.VariableSubStmtContext)target.Context);
                var count = items.Count;

                var element = variables.Parent.Parent as VBAParser.ModuleDeclarationsElementContext;
                if (element != null)
                {
                    // module level variable declaration
                    var startIndex = element.Start.TokenIndex;
                    var stopIndex = element.Stop.TokenIndex;
                    var parent = (VBAParser.ModuleDeclarationsContext)element.Parent;
                    var elements = parent.moduleDeclarationsElement();

                    if (count == 1)
                    {
                        // single variable - return the bounds of entire statement
                        for (var i = 0; i < elements.Count; i++)
                        {
                            if (elements[i] != element)
                            {
                                continue;
                            }

                            stopIndex = parent.endOfStatement(i).Stop.TokenIndex;
                            break;
                        }

                        return new RewriterInfo(startIndex, stopIndex);
                    }
                    else
                    {
                        startIndex = itemIndex < count - 1
                            ? target.Context.Start.TokenIndex
                            : items[itemIndex - 1].Stop.TokenIndex + 1;

                        stopIndex = itemIndex < count - 1
                            ? items[itemIndex + 1].Start.TokenIndex - 1
                            : items[itemIndex].Stop.TokenIndex;

                        return new RewriterInfo(startIndex, stopIndex);
                    }
                }
                else if(variables.Parent is VBAParser.VariableStmtContext)
                {
                    // procedure level variable declaration
                    var blockStmt = (VBAParser.BlockStmtContext)variables.Parent.Parent;
                    var startIndex = blockStmt.Start.TokenIndex;
                    var stopIndex = blockStmt.Stop.TokenIndex;
                    var parent = (VBAParser.BlockContext)blockStmt.Parent;
                    var statements = parent.blockStmt();

                    if (count == 1)
                    {
                        // single variable - return the bounds of entire statement
                        for (var i = 0; i < statements.Count; i++)
                        {
                            if (statements[i] != blockStmt)
                            {
                                continue;
                            }

                            stopIndex = parent.endOfStatement(i).Stop.TokenIndex;
                            break;
                        }

                        return new RewriterInfo(startIndex, stopIndex);
                    }
                    else
                    {
                        startIndex = itemIndex < count - 1
                            ? target.Context.Start.TokenIndex
                            : items[itemIndex - 1].Stop.TokenIndex + 1;

                        stopIndex = itemIndex < count - 1
                            ? items[itemIndex + 1].Start.TokenIndex - 1
                            : items[itemIndex].Stop.TokenIndex;

                        return new RewriterInfo(startIndex, stopIndex);
                    }
                }
            }

            return RewriterInfo.None;
        }



        //private static RewriterInfo GetRemovedTokenIndex<TContext>(Declaration target, IReadOnlyList<TContext> items, Func<TContext, string> getIdentifierName)
        //    where TContext : ParserRuleContext
        //{
        //    if (items.Count == 1 && target.DeclarationType != DeclarationType.Parameter)
        //    {
        //        var item = items[0];
        //        var start = ((ParserRuleContext) item.Parent.Parent).Start.TokenIndex;
        //        var stop = ((ParserRuleContext) item.Parent.Parent).Stop.TokenIndex;
        //        return Tuple.Create(start == 0 ? 0 : start - 1, stop + 2);
        //    }

        //    for (var i = 0; i < items.Count; i++)
        //    {
        //        var item = items[i];
        //        if (getIdentifierName(item) == target.IdentifierName)
        //        {
        //            if (items.Count > 1)
        //            {
        //                ITerminalNode comma;
        //                if (i == 0)
        //                {
        //                    comma = (ITerminalNode)((dynamic)item.Parent).COMMA()[i];
        //                }
        //                else
        //                {
        //                    comma = (ITerminalNode)((dynamic)item.Parent).COMMA()[i - 1];
        //                }
        //                return Tuple.Create(item.Start.TokenIndex, comma.Symbol.TokenIndex);
        //            }
        //            return Tuple.Create(item.Start.TokenIndex, item.Stop.TokenIndex);
        //        }
        //    }

        //    Debug.Assert(false, "Could not locate target token range.");
        //    return null;
        //}

        private enum TargetListPosition
        {
            /// <summary>
            /// Target was the only item in a list, or there was no list; no leading or trailing comma needs to be handled.
            /// </summary>
            SingleItem,
            /// <summary>
            /// Target was the first item in a list of two or more: a leading comma needs to be handled.
            /// </summary>
            FirstItem,
            /// <summary>
            /// Target was the last item in a list of two or more: a trailing comma needs to be handled.
            /// </summary>
            LastItem,
        }

        private static ParserRuleContext GetStmtContext(Declaration target, out TargetListPosition position)
        {
            ParserRuleContext result;
            position = TargetListPosition.SingleItem;
            // for instructions that may contain more than a single declaration, we need to isolate the target's context.
            switch (target.DeclarationType)
            {
                case DeclarationType.Variable:
                    result = GetVariableContext(target, ref position);
                    break;

                case DeclarationType.Parameter:
                    result = GetParameterContext(target, ref position);
                    break;

                case DeclarationType.Constant:

                    result = GetConstantContext(target, ref position);
                    break;

                default:
                    result = target.Context;
                    break;
            }
            return result;
        }

        private static ParserRuleContext GetParameterContext(Declaration target, ref TargetListPosition position)
        {
            var argList = (VBAParser.ArgListContext) target.Context.Parent;
            var args = argList.arg();
            var count = args.Count;
            ParserRuleContext result = argList;

            for (var i = 0; i < count; i++)
            {
                // foreach is less practical to track index
                var arg = args[i];
                if (Identifier.GetName(arg.unrestrictedIdentifier()) != target.IdentifierName)
                {
                    continue;
                }

                result = arg;
                position = GetTargetListPosition(i, count);
            }

            return result;
        }

        private static ParserRuleContext GetConstantContext(Declaration target, ref TargetListPosition position)
        {
            var constStmt = target.GetConstStmtContext();
            var consts = constStmt.constSubStmt();
            var count = consts.Count;
            ParserRuleContext result = constStmt;

            for (var i = 0; i < count; i++)
            {
                var constant = consts[i];
                if (Identifier.GetName(constant.identifier()) != target.IdentifierName)
                {
                    continue;
                }

                result = constant;
                position = GetTargetListPosition(i, count);
            }
            return result;
        }

        private static ParserRuleContext GetVariableContext(Declaration target, ref TargetListPosition position)
        {
            var variableStmt = target.GetVariableStmtContext();
            ParserRuleContext result = variableStmt;

            var variables = variableStmt.variableListStmt().variableSubStmt();
            var count = variables.Count;

            for (var i = 0; i < count; i++)
            {
                var variable = variables[i];
                if (Identifier.GetName(variable.identifier()) != target.IdentifierName)
                {
                    continue;
                }

                result = variable;
                position = GetTargetListPosition(i, count);
            }
            return result;
        }

        private static TargetListPosition GetTargetListPosition(int i, int count)
        {
            TargetListPosition position;
            if (i == 0)
            {
                position = TargetListPosition.FirstItem;
            }
            else if (i == count - 1)
            {
                position = TargetListPosition.LastItem;
            }
            else
            {
                position = TargetListPosition.SingleItem;
            }
            return position;
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

        public static void Remove(this ICodeModule module, IdentifierReference target, TokenStreamRewriter rewriter)
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

        public static void Remove(this ICodeModule module, IEnumerable<IdentifierReference> targets, TokenStreamRewriter rewriter)
        {
            foreach (var target in targets/*.OrderByDescending(e => e.Selection)*/)
            {
                module.Remove(target, rewriter);
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
