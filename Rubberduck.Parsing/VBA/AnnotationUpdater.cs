using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Antlr4.Runtime.Tree;
using NLog;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA
{
    public class AnnotationUpdater : IAnnotationUpdater
    {
        private readonly IParseTreeProvider _parseTreeProvider; 

        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public AnnotationUpdater(IParseTreeProvider parseTreeProvider)
        {
            _parseTreeProvider = parseTreeProvider;
        }

        public void AddAnnotation(IRewriteSession rewriteSession, QualifiedContext context, IAnnotation annotationInfo, IReadOnlyList<string> values = null)
        {
            var annotationValues = values ?? new List<string>();

            if (context == null)
            {
                _logger.Warn("Tried to add an annotation to a context that is null.");
                _logger.Trace($"Tried to add annotation {annotationInfo.Name} with values {AnnotationValuesText(annotationValues)} to a context that is null.");
                return;
            }

            if (rewriteSession.TargetCodeKind != CodeKind.CodePaneCode)
            {
                _logger.Warn($"Tried to add an annotation with a rewriter not suitable to annotate contexts. (target code kind = {rewriteSession.TargetCodeKind})");
                _logger.Trace($"Tried to add annotation {annotationInfo.Name} with values {AnnotationValuesText(annotationValues)} to {context.Context.GetText()} at {context.Context.GetSelection()} in module {context.ModuleName} using a rewriter not suitable for annotations.");
                return;
            }

            AddAnnotation(rewriteSession, context.ModuleName, context.Context, annotationInfo, annotationValues);
        }

        private void AddAnnotation(IRewriteSession rewriteSession, QualifiedModuleName moduleName, ParserRuleContext context, IAnnotation annotationInfo, IReadOnlyList<string> values = null)
        {
            var annotationValues = values ?? new List<string>();

            if (context == null)
            {
                _logger.Warn("Tried to add an annotation to a context that is null.");
                _logger.Trace($"Tried to add annotation {annotationInfo.Name} with values {AnnotationValuesText(annotationValues)} to a context that is null.");
                return;
            }

            var annotationText = AnnotationText(annotationInfo.Name, annotationValues);

            int? startOfLogicalLine = _parseTreeProvider
                .GetLogicalLines(moduleName)
                ?.StartOfContainingLogicalLine(context.start.Line);

            if (!startOfLogicalLine.HasValue)
            {
                _logger.Warn("Tried to add an annotation to a context that is not on any known logical line of the module.");
                _logger.Trace($"Tried to add annotation {annotationInfo.Name} with values {AnnotationValuesText(annotationValues)} to a context with startline {context.start.Line} in module {moduleName} that is outside all known logical lines.");
                return;
            }

            if (startOfLogicalLine.Value == 1)
            {
                InsertAtStartOfModule(rewriteSession, moduleName, annotationText);
                return;
            }

            var endOfLineBeforeLogicalLine = EndOfLineBeforePhysicalLine(context, startOfLogicalLine.Value);

            var codeToAdd = endOfLineBeforeLogicalLine.TryGetFollowingContext(out VBAParser.WhiteSpaceContext whitespaceAtStartOfLine)
                            ? $"{whitespaceAtStartOfLine.GetText()}{annotationText}{Environment.NewLine}"
                            : $"{annotationText}{Environment.NewLine}";
            var rewriter = rewriteSession.CheckOutModuleRewriter(moduleName);
            rewriter.InsertAfter(endOfLineBeforeLogicalLine.stop.TokenIndex, codeToAdd);
        }

        private static void InsertAtStartOfModule(IRewriteSession rewriteSession, QualifiedModuleName moduleName, string annotationText)
        {
            var codeToAdd = $"{annotationText}{Environment.NewLine}";
            var rewriter = rewriteSession.CheckOutModuleRewriter(moduleName);
            rewriter.InsertBefore(0, codeToAdd);
        }

        private static string AnnotationText(IAnnotation annotationInformation, IReadOnlyList<string> values)
        {
            return AnnotationText(annotationInformation.Name, values);
        }

        private static string AnnotationText(string annotationType, IReadOnlyList<string> values)
        {
            return $"'{ParseTreeAnnotation.ANNOTATION_MARKER}{AnnotationBaseText(annotationType, values)}";
        }

        private static string AnnotationBaseText(string annotationType, IReadOnlyList<string> values)
        {
            return $"{annotationType}{(values.Any() ? $" {AnnotationValuesText(values)}" : string.Empty)}";
        }

        private static string AnnotationValuesText(IEnumerable<string> annotationValues)
        {
            return string.Join(", ", annotationValues);
        }

        private static VBAParser.EndOfLineContext EndOfLineBeforePhysicalLine(ParserRuleContext context, int physicalLine)
        {
            var moduleContext = context.GetAncestor<VBAParser.ModuleContext>();
            var endOfLineListener = new EndOfLineListener();
            ParseTreeWalker.Default.Walk(endOfLineListener, moduleContext);
            var previousEol = endOfLineListener.Contexts
                .OrderBy(eol => eol.Start.TokenIndex)
                .LastOrDefault(eol => eol.start.Line < physicalLine);
            return previousEol;
        }

        public void AddAnnotation(IRewriteSession rewriteSession, Declaration declaration, IAnnotation annotationInfo, IReadOnlyList<string> values = null)
        {
            var annotationValues = values ?? new List<string>();

            if (declaration == null)
            {
                _logger.Warn("Tried to add an annotation to a declaration that is null.");
                _logger.Trace($"Tried to add annotation {annotationInfo.Name} with values {AnnotationValuesText(annotationValues)} to a declaration that is null.");
                return;
            }

            if (declaration.DeclarationType.HasFlag(DeclarationType.Module))
            {
                AddModuleAnnotation(rewriteSession, declaration, annotationInfo, annotationValues);
            }
            else if (declaration.DeclarationType.HasFlag(DeclarationType.Variable))
            {
                AddVariableAnnotation(rewriteSession, declaration, annotationInfo, annotationValues);
            }
            else
            {
                AddMemberAnnotation(rewriteSession, declaration, annotationInfo, annotationValues);
            }
        }

        private void AddModuleAnnotation(IRewriteSession rewriteSession, Declaration declaration, IAnnotation annotationInfo, IReadOnlyList<string> annotationValues)
        {
            if (!annotationInfo.Target.HasFlag(AnnotationTarget.Module))
            {
                _logger.Warn("Tried to add an annotation without the module annotation flag to a module.");
                _logger.Trace($"Tried to add the annotation {annotationInfo.Name} with values {AnnotationValuesText(annotationValues)} to the module {declaration.QualifiedModuleName}.");
                return;
            }

            if (rewriteSession.TargetCodeKind != CodeKind.CodePaneCode && rewriteSession.TargetCodeKind != CodeKind.AttributesCode)
            {
                _logger.Warn($"Tried to add an annotation to a module with a rewriter not suitable for annotations. (target code kind = {rewriteSession.TargetCodeKind})");
                _logger.Trace($"Tried to add annotation {annotationInfo.Name} with values {AnnotationValuesText(annotationValues)} to the module {declaration.QualifiedModuleName} using a rewriter not suitable for annotations.");
                return;
            }

            var codeToAdd = AnnotationText(annotationInfo, annotationValues);

            var rewriter = rewriteSession.CheckOutModuleRewriter(declaration.QualifiedModuleName);

            if (rewriteSession.TargetCodeKind == CodeKind.AttributesCode)
            {
                InsertAfterLastModuleAttribute(rewriter, declaration.QualifiedModuleName, codeToAdd);
            }
            else
            {
                var codeToInsert = codeToAdd + Environment.NewLine;
                rewriter.InsertBefore(0, codeToInsert);
            }
        }

        private void InsertAfterLastModuleAttribute(IModuleRewriter rewriter, QualifiedModuleName module, string codeToAdd)
        {
            var moduleParseTree = (ParserRuleContext)_parseTreeProvider.GetParseTree(module, CodeKind.AttributesCode);
            var lastModuleAttribute = moduleParseTree.GetDescendents<VBAParser.ModuleAttributesContext>()
                .Where(moduleAttributes => moduleAttributes.attributeStmt() != null)
                .SelectMany(moduleAttributes => moduleAttributes.attributeStmt())
                .OrderBy(moduleAttribute => moduleAttribute.stop.TokenIndex)
                .LastOrDefault();
            if (lastModuleAttribute == null)
            {
                //This should never happen for a real module.
                var codeToInsert = codeToAdd + Environment.NewLine;
                rewriter.InsertBefore(0, codeToInsert);
            }
            else
            {
                var codeToInsert = Environment.NewLine + codeToAdd;
                rewriter.InsertAfter(lastModuleAttribute.stop.TokenIndex, codeToInsert);
            }
        }

        private void AddVariableAnnotation(IRewriteSession rewriteSession, Declaration declaration, IAnnotation annotationInfo, IReadOnlyList<string> annotationValues)
        {
            if (!annotationInfo.Target.HasFlag(AnnotationTarget.Variable))
            {
                _logger.Warn("Tried to add an annotation without the variable annotation flag to a variable declaration.");
                _logger.Trace($"Tried to add the annotation {annotationInfo.Name} with values {AnnotationValuesText(annotationValues)} to the variable declaration for {declaration.QualifiedName}.");
                return;
            }

            if (rewriteSession.TargetCodeKind != CodeKind.CodePaneCode && (rewriteSession.TargetCodeKind != CodeKind.AttributesCode || declaration.AttributesPassContext == null))
            {
                _logger.Warn($"Tried to add an annotation to a variable with a rewriter not suitable for annotations to the variable. (target code kind = {rewriteSession.TargetCodeKind})");
                _logger.Trace($"Tried to add annotation {annotationInfo.Name} with values {AnnotationValuesText(annotationValues)} to the the variable {declaration.IdentifierName} at {declaration.Selection} in module {declaration.QualifiedModuleName} using a rewriter not suitable for annotations.");
                return;
            }

            var context = rewriteSession.TargetCodeKind == CodeKind.CodePaneCode
                ? declaration.Context
                : declaration.AttributesPassContext;

            AddAnnotation(rewriteSession, declaration.QualifiedModuleName, context, annotationInfo, annotationValues);
        }

        private void AddMemberAnnotation(IRewriteSession rewriteSession, Declaration declaration, IAnnotation annotationInfo, IReadOnlyList<string> annotationValues)
        {
            if (!annotationInfo.Target.HasFlag(AnnotationTarget.Member))
            {
                _logger.Warn("Tried to add an annotation without the member annotation flag to a member declaration.");
                _logger.Trace($"Tried to add the annotation {annotationInfo.Name} with values {AnnotationValuesText(annotationValues)} to the member declaration for {declaration.QualifiedName}.");
                return;
            }

            if (rewriteSession.TargetCodeKind != CodeKind.CodePaneCode && (rewriteSession.TargetCodeKind != CodeKind.AttributesCode || declaration.AttributesPassContext == null))
            {
                _logger.Warn($"Tried to add an annotation to a member with a rewriter not suitable for annotations to the member. (target code kind = {rewriteSession.TargetCodeKind})");
                _logger.Trace($"Tried to add annotation {annotationInfo.Name} with values {AnnotationValuesText(annotationValues)} to the the member {declaration.IdentifierName} at {declaration.Selection} in module {declaration.QualifiedModuleName} using a rewriter not suitable for annotations.");
                return;
            }

            var context = rewriteSession.TargetCodeKind == CodeKind.CodePaneCode
                ? declaration.Context
                : declaration.AttributesPassContext;

            AddAnnotation(rewriteSession, declaration.QualifiedModuleName, context, annotationInfo, annotationValues);
        }

        public void AddAnnotation(IRewriteSession rewriteSession, IdentifierReference reference, IAnnotation annotationInfo, IReadOnlyList<string> values = null)
        {
            var annotationValues = values ?? new List<string>();

            if (reference == null)
            {
                _logger.Warn("Tried to add an annotation to an identifier reference that is null.");
                _logger.Trace($"Tried to add annotation {annotationInfo.Name} with values {AnnotationValuesText(annotationValues)} to an identifier reference that is null.");
                return;
            }

            if (!annotationInfo.Target.HasFlag(AnnotationTarget.Identifier))
            {
                _logger.Warn("Tried to add an annotation without the identifier reference annotation flag to an identifier reference.");
                _logger.Trace($"Tried to add annotation {annotationInfo.Name} with values {AnnotationValuesText(annotationValues)} to the identifier reference to {reference.Declaration.QualifiedName} at {reference.Selection} in module {reference.QualifiedModuleName}.");
                return;
            }

            if (rewriteSession.TargetCodeKind != CodeKind.CodePaneCode)
            {
                _logger.Warn($"Tried to add an annotation to an identifier reference with a rewriter not suitable for annotations to references. (target code kind = {rewriteSession.TargetCodeKind})");
                _logger.Trace($"Tried to add annotation {annotationInfo.Name} with values {AnnotationValuesText(annotationValues)} to the the identifier reference {reference.IdentifierName} at {reference.Selection} in module {reference.QualifiedModuleName} using a rewriter not suitable for annotations.");
                return;
            }

            AddAnnotation(rewriteSession, new QualifiedContext(reference.QualifiedModuleName, reference.Context), annotationInfo, annotationValues);
        }

        public void RemoveAnnotation(IRewriteSession rewriteSession, IParseTreeAnnotation annotation)
        {
            if (annotation == null)
            {
                _logger.Warn("Tried to remove an annotation that is null.");
                return;
            }

            if (rewriteSession.TargetCodeKind != CodeKind.CodePaneCode)
            {
                _logger.Warn($"Tried to remove an annotation with a rewriter not suitable for annotationss. (target code kind = {rewriteSession.TargetCodeKind})");
                _logger.Trace($"Tried to remove annotation {annotation.Annotation.Name} at {annotation.QualifiedSelection.Selection} in module {annotation.QualifiedSelection.QualifiedName} using a rewriter not suitable for annotations.");
                return;
            }

            var annotationContext = annotation.Context;
            var annotationList = (VBAParser.AnnotationListContext)annotationContext.Parent;

            var rewriter = rewriteSession.CheckOutModuleRewriter(annotation.QualifiedSelection.QualifiedName);

            var annotations = annotationList.annotation();
            if (annotations.Length == 1)
            {
                RemoveSingleAnnotation(rewriter, annotationContext, annotationList);
            }

            RemoveAnnotationMarker(rewriter, annotationContext);
            rewriter.Remove(annotationContext);
        }

        private static void RemoveSingleAnnotation(IModuleRewriter rewriter, VBAParser.AnnotationContext annotationContext, VBAParser.AnnotationListContext annotationListContext)
        {
            var commentSeparator = annotationListContext.COLON();
            if(commentSeparator == null)
            {
                RemoveEntireLine(rewriter, annotationContext);
            }
            else
            {
                RemoveAnnotationMarker(rewriter, annotationContext);
                rewriter.Remove(annotationContext);
                rewriter.Remove(commentSeparator);
            }
        }

        private static void RemoveEntireLine(IModuleRewriter rewriter, ParserRuleContext contextInCommentOrAnnotation)
        {
            var previousEndOfLineContext = EndOfLineBeforePhysicalLine(contextInCommentOrAnnotation, contextInCommentOrAnnotation.start.Line);
            var containingCommentOrAnnotationContext = contextInCommentOrAnnotation.GetAncestor<VBAParser.CommentOrAnnotationContext>();

            if (previousEndOfLineContext == null)
            {
                //We are on the first logical line.
                rewriter.RemoveRange(0, containingCommentOrAnnotationContext.stop.TokenIndex);
            }
            else if (containingCommentOrAnnotationContext.Eof() != null)
            {
                //We are on the last logical line. So swallow the NEWLINE from the previous end of line.
                rewriter.RemoveRange(previousEndOfLineContext.stop.TokenIndex, containingCommentOrAnnotationContext.stop.TokenIndex);
            }
            else
            {
                rewriter.RemoveRange(previousEndOfLineContext.stop.TokenIndex + 1, containingCommentOrAnnotationContext.stop.TokenIndex);
            }
        }

        private static void RemoveAnnotationMarker(IModuleRewriter rewriter, VBAParser.AnnotationContext annotationContext)
        {
            var endOfAnnotationMarker = annotationContext.start.TokenIndex - 1;
            var startOfAnnotationMarker = endOfAnnotationMarker - ParseTreeAnnotation.ANNOTATION_MARKER.Length + 1;
            rewriter.RemoveRange(startOfAnnotationMarker, endOfAnnotationMarker);
        }

        public void RemoveAnnotations(IRewriteSession rewriteSession, IEnumerable<IParseTreeAnnotation> annotations)
        {
            if (annotations == null)
            {
                return;
            }

            if (rewriteSession.TargetCodeKind != CodeKind.CodePaneCode && rewriteSession.TargetCodeKind != CodeKind.AttributesCode)
            {
                _logger.Warn($"Tried to remove multiple annotations with a rewriter not suitable for annotations. (target code kind = {rewriteSession.TargetCodeKind})");
                return;
            }

            var annotationsByAnnotationList = annotations.Distinct()
                .GroupBy(annotation => new QualifiedContext(annotation.QualifiedSelection.QualifiedName, (ParserRuleContext)annotation.Context.Parent))
                .ToDictionary(grouping => grouping.Key, grouping => grouping.ToList());

            if (!annotationsByAnnotationList.Keys.Any())
            {
                return;
            }

            foreach (var qualifiedAnnotationList in annotationsByAnnotationList.Keys)
            {
                var annotationList = (VBAParser.AnnotationListContext) qualifiedAnnotationList.Context;
                if (annotationList.commentBody() == null && annotationList.annotation().Length == annotationsByAnnotationList[qualifiedAnnotationList].Count)
                {
                    //We want to remove all annotations in the list. So, we remove the entire line.
                    //This does not really work if there are multiple consecutive lines at the end of the file that need to be removed,
                    //but I think we can live with leaving an empty line in this edge-case.
                    var rewriter = rewriteSession.CheckOutModuleRewriter(qualifiedAnnotationList.ModuleName);
                    RemoveEntireLine(rewriter, annotationList);
                }
                else
                {
                    foreach (var annotation in annotationsByAnnotationList[qualifiedAnnotationList])
                    {
                        RemoveAnnotation(rewriteSession, annotation);
                    }
                }
            }
        }

        public void UpdateAnnotation(IRewriteSession rewriteSession, IParseTreeAnnotation annotation, IAnnotation annotationInfo, IReadOnlyList<string> newValues = null)
        {
            var newAnnotationValues = newValues ?? new List<string>();

            if (annotation == null)
            {
                _logger.Warn("Tried to replace an annotation that is null.");
                _logger.Trace($"Tried to replace an annotation that is null with an annotation {annotationInfo.Name} with values {AnnotationValuesText(newAnnotationValues)}.");
                return;
            }

            if (rewriteSession.TargetCodeKind != CodeKind.CodePaneCode && rewriteSession.TargetCodeKind != CodeKind.AttributesCode)
            {
                _logger.Warn($"Tried to update an annotation with a rewriter not suitable for annotationss. (target code kind = {rewriteSession.TargetCodeKind})");
                _logger.Trace($"Tried to update annotation {annotation.Annotation.Name} at {annotation.QualifiedSelection.Selection} in module {annotation.QualifiedSelection.QualifiedName} with annotation {annotationInfo.Name} with values {AnnotationValuesText(newAnnotationValues)} using a rewriter not suitable for annotations.");
                return;
            }

            //If there are no common flags, the annotations cannot apply to the same target.
            if ((annotation.Annotation.Target & annotationInfo.Target) == 0)
            {
                _logger.Warn("Tried to replace an annotation with an annotation without common flags.");
                _logger.Trace($"Tried to replace an annotation {annotation.Annotation.Name} with values {AnnotationValuesText(newValues)} at {annotation.QualifiedSelection.Selection} in module {annotation.QualifiedSelection.QualifiedName} with an annotation {annotationInfo.Name} with values {AnnotationValuesText(newAnnotationValues)}, which does not have any common flags.");
                return;
            }
            
            var context = annotation.Context;
            var whitespaceAtEnd = context.whiteSpace()?.GetText() ?? string.Empty;
            var codeReplacement = $"{AnnotationBaseText(annotationInfo.Name, newAnnotationValues)}{whitespaceAtEnd}";

            var rewriter = rewriteSession.CheckOutModuleRewriter(annotation.QualifiedSelection.QualifiedName);
            rewriter.Replace(annotation.Context, codeReplacement);
        }

        private class EndOfLineListener : VBAParserBaseListener
        {
            private readonly IList<VBAParser.EndOfLineContext> _contexts = new List<VBAParser.EndOfLineContext>();
            public IEnumerable<VBAParser.EndOfLineContext> Contexts => _contexts;

            public override void ExitEndOfLine([NotNull] VBAParser.EndOfLineContext context)
            {
                _contexts.Add(context);
            }
        }
    }
}