using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Antlr4.Runtime.Tree;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Refactorings.ExtractMethod
{
    public class ExtractMethodSelectionValidation : IExtractMethodSelectionValidation
    {
        private readonly ICodeModule _codeModule;
        private readonly IEnumerable<Declaration> _declarations;
        private List<Tuple<ParserRuleContext, string>> _invalidContexts = new List<Tuple<ParserRuleContext, string>>();
        private List<VBAParser.BlockStmtContext> _finalResults = new List<VBAParser.BlockStmtContext>();

        public ExtractMethodSelectionValidation(IEnumerable<Declaration> declarations, ICodeModule codeModule)
        {
            _declarations = declarations;
            _codeModule = codeModule;
        }

        public IEnumerable<Tuple<ParserRuleContext, string>> InvalidContexts => _invalidContexts;

        public IEnumerable<VBAParser.BlockStmtContext> SelectedContexts => _finalResults;

        public bool ValidateSelection(QualifiedSelection qualifiedSelection)
        {
            var selection = qualifiedSelection.Selection;
            var procedures = _declarations.Where(d => d.ComponentName == qualifiedSelection.QualifiedName.ComponentName && d.IsUserDefined && (DeclarationExtensions.ProcedureTypes.Contains(d.DeclarationType)));
            var declarations = procedures as IList<Declaration> ?? procedures.ToList();
            Declaration ProcOfLine(int sl) => declarations.FirstOrDefault(d => d.Context.Start.Line < sl && d.Context.Stop.EndLine() > sl);

            var startLine = selection.StartLine;
            var endLine = selection.EndLine;

            // End of line is easy
            var procEnd = ProcOfLine(endLine);
            if (procEnd == null)
            {
                return false;
            }

            var procStart = ProcOfLine(startLine);
            if (procStart == null)
            {
                return false;
            }

            var procStartContext = procStart.Context;
            VBAParser.EndOfStatementContext procEndOfSignature;

            switch (procStartContext)
            {
                case VBAParser.FunctionStmtContext funcStmt:
                    procEndOfSignature = funcStmt.endOfStatement();
                    break;
                case VBAParser.SubStmtContext subStmt:
                    procEndOfSignature = subStmt.endOfStatement();
                    break;
                case VBAParser.PropertyGetStmtContext getStmt:
                    procEndOfSignature = getStmt.endOfStatement();
                    break;
                case VBAParser.PropertyLetStmtContext letStmt:
                    procEndOfSignature = letStmt.endOfStatement();
                    break;
                case VBAParser.PropertySetStmtContext setStmt:
                    procEndOfSignature = setStmt.endOfStatement();
                    break;
                default:
                    return false;
            }
            
            if (!(procEnd.QualifiedSelection.Equals(procStart.QualifiedSelection)
                && (procEndOfSignature.Start.Line < selection.StartLine
                || procEndOfSignature.Start.Line == selection.StartLine && procEndOfSignature.Start.Column < selection.StartColumn)
                ))
                return false;

            
            /* At this point, we know the selection is within a single procedure. We need to validate that the user's
             * selection in fact contain only BlockStmt and not other stuff that might not be so extractable.
             */
            var visitor = new ExtractValidatorVisitor(qualifiedSelection, _invalidContexts);
            var results = visitor.Visit(procStartContext);
            _invalidContexts = visitor.InvalidContexts;

            if (!_invalidContexts.Any())
            {
                if (_codeModule.ContainsCompilationDirectives(selection))
                {
                    return false;
                }
                // We've proved that there are no invalid statements contained in the selection. However, we need to analyze
                // the statements to ensure they are not partial selections.

                // The visitor will not return the results in a sorted manner, so we need to arrange the contexts in the same order.
                var blockStmtContexts = results as IList<VBAParser.BlockStmtContext> ?? results.ToList();
                var sorted = blockStmtContexts.OrderBy(context => context.Start.StartIndex);
                ContextIsContainedOnce(sorted, ref _finalResults, qualifiedSelection);
                return blockStmtContexts.Any() && !_invalidContexts.Any() && _finalResults.Any();
            }
            return false;
        }

        /// <summary>
        /// The function ensure that we return only top-level BlockStmtContexts that
        /// exists within an user's selection, excluding any nested BlockStmtContexts
        /// which are also "selected" and thus ensure that we build an unique list 
        /// of BlockStmtContexts that corresponds to the user's selection. The function
        /// also will validate there are no overlapping selections which could be invalid.
        /// </summary>
        /// <param name="sortedResults">The context to test</param>
        /// <param name="aggregate">The list of contexts we already added to verify we are not adding one of its children or itself more than once</param>
        /// <param name="qualifiedSelection"></param>
        /// <returns>Boolean with true indicating that it's the first time we encountered a context in a user's selection and we can safely add it to the list</returns>
        private void ContextIsContainedOnce(IEnumerable<VBAParser.BlockStmtContext> sortedResults, ref List<VBAParser.BlockStmtContext> aggregate, QualifiedSelection qualifiedSelection)
        {
            foreach (var context in sortedResults)
            {
                if (qualifiedSelection.Selection.Contains(context))
                {
                    if (!aggregate.Any(otherContext => otherContext.GetSelection().Contains(context) && context != otherContext))
                    {
                        aggregate.Add(context);
                    }
                }
                else
                {
                    // We need to check if there was a partial selection made which would be invalid. It's OK if it's wholly contained inside
                    // a context (e.g. an inner If/End If block within a bigger If/End If was selected which is legal. However, selecting only
                    // part of inner If/End If block and a part of the outermost If/End If block should be illegal).
                    if (qualifiedSelection.Selection.Overlaps(context.GetSelection()) && !qualifiedSelection.Selection.IsContainedIn(context))
                    {
                        _invalidContexts.Add(Tuple.Create(context as ParserRuleContext, "Extract method must contain selection that represents a set of complete statements. It cannot extract a part of statement."));
                    }
                }
            }
        }

        private class ExtractValidatorVisitor : VBAParserBaseVisitor<IEnumerable<VBAParser.BlockStmtContext>>
        {
            private readonly QualifiedSelection _qualifiedSelection;

            public ExtractValidatorVisitor(QualifiedSelection qualifiedSelection, List<Tuple<ParserRuleContext, string>> invalidContexts)
            {
                _qualifiedSelection = qualifiedSelection;
                InvalidContexts = invalidContexts;
            }

            public List<Tuple<ParserRuleContext, string>> InvalidContexts { get; }

            protected override IEnumerable<VBAParser.BlockStmtContext> DefaultResult => new List<VBAParser.BlockStmtContext>();
            
            public override IEnumerable<VBAParser.BlockStmtContext> VisitBlockStmt([NotNull] VBAParser.BlockStmtContext context)
            {
                var children = base.VisitBlockStmt(context);
                return InvalidContexts.Count == 0 ? children.Concat(new List<VBAParser.BlockStmtContext> { context }) : null;
            }

            protected override IEnumerable<VBAParser.BlockStmtContext> AggregateResult(IEnumerable<VBAParser.BlockStmtContext> aggregate, IEnumerable<VBAParser.BlockStmtContext> nextResult)
            {
                return InvalidContexts.Count == 0 ? aggregate.Concat(nextResult) : null;
            }

            protected override bool ShouldVisitNextChild(IRuleNode node, IEnumerable<VBAParser.BlockStmtContext> currentResult)
            {
                // Don't visit any more children if we have any invalid contexts
                return (InvalidContexts.Count == 0);
            }

            public override IEnumerable<VBAParser.BlockStmtContext> VisitErrorStmt([NotNull] VBAParser.ErrorStmtContext context)
            {
                if (_qualifiedSelection.Selection.Contains(context))
                {
                    InvalidContexts.Add(Tuple.Create(context as ParserRuleContext, "Extract method cannot extract methods that contains an Error statement."));
                    return null;
                }

                return base.VisitErrorStmt(context);
            }

            public override IEnumerable<VBAParser.BlockStmtContext> VisitEndStmt([NotNull] VBAParser.EndStmtContext context)
            {
                if (_qualifiedSelection.Selection.Contains(context))
                {
                    InvalidContexts.Add(Tuple.Create(context as ParserRuleContext, "Extract method cannot extract methods that contains an End statement."));
                    return null;
                }

                return base.VisitEndStmt(context);
            }

            public override IEnumerable<VBAParser.BlockStmtContext> VisitExitStmt([NotNull] VBAParser.ExitStmtContext context)
            {
                if (_qualifiedSelection.Selection.Contains(context))
                {
                    InvalidContexts.Add(Tuple.Create(context as ParserRuleContext, "Extract method cannot extract methods that contains an Exit statement"));
                    return null;
                }

                return base.VisitExitStmt(context);
            }

            public override IEnumerable<VBAParser.BlockStmtContext> VisitGoSubStmt([NotNull] VBAParser.GoSubStmtContext context)
            {
                if (_qualifiedSelection.Selection.Contains(context))
                {
                    InvalidContexts.Add(Tuple.Create(context as ParserRuleContext, "Extract method cannot extract methods that contains a GoSub statement"));
                    return null;
                }

                return base.VisitGoSubStmt(context);
            }

            public override IEnumerable<VBAParser.BlockStmtContext> VisitGoToStmt([NotNull] VBAParser.GoToStmtContext context)
            {
                if (_qualifiedSelection.Selection.Contains(context))
                {
                    InvalidContexts.Add(Tuple.Create(context as ParserRuleContext, "Extract method cannot extract methods that contains a GoTo statement"));
                    return null;
                }

                return base.VisitGoToStmt(context);
            }

            public override IEnumerable<VBAParser.BlockStmtContext> VisitOnErrorStmt([NotNull] VBAParser.OnErrorStmtContext context)
            {
                if (_qualifiedSelection.Selection.Contains(context))
                {
                    InvalidContexts.Add(Tuple.Create(context as ParserRuleContext, "Extract method cannot extract methods that contains a On Error statement"));
                    return null;
                }

                return base.VisitOnErrorStmt(context);
            }
            
            public override IEnumerable<VBAParser.BlockStmtContext> VisitIdentifierStatementLabel([NotNull] VBAParser.IdentifierStatementLabelContext context)
            {
                if (_qualifiedSelection.Selection.Contains(context))
                {
                    InvalidContexts.Add(Tuple.Create(context as ParserRuleContext, "Extract method cannot extract methods that contains a Line Label statement"));
                    return null;
                }

                return base.VisitIdentifierStatementLabel(context);
            }

            public override IEnumerable<VBAParser.BlockStmtContext> VisitCombinedLabels([NotNull] VBAParser.CombinedLabelsContext context)
            {
                if (_qualifiedSelection.Selection.Contains(context))
                {
                    InvalidContexts.Add(Tuple.Create(context as ParserRuleContext, "Extract method cannot extract methods that contains a Line Label statement"));
                    return base.VisitCombinedLabels(context);
                }

                return VisitCombinedLabels(context);
            }

            public override IEnumerable<VBAParser.BlockStmtContext> VisitOnGoSubStmt([NotNull] VBAParser.OnGoSubStmtContext context)
            {
                if (_qualifiedSelection.Selection.Contains(context))
                {
                    InvalidContexts.Add(Tuple.Create(context as ParserRuleContext, "Extract method cannot extract methods that contains a On ... GoSub statement"));
                    return null;
                }

                return VisitOnGoSubStmt(context);
            }

            public override IEnumerable<VBAParser.BlockStmtContext> VisitOnGoToStmt([NotNull] VBAParser.OnGoToStmtContext context)
            {
                if (_qualifiedSelection.Selection.Contains(context))
                {
                    InvalidContexts.Add(Tuple.Create(context as ParserRuleContext, "Extract method cannot extract methods that contains a On ... GoTo statement"));
                    return null;
                }

                return VisitOnGoToStmt(context);
            }

            public override IEnumerable<VBAParser.BlockStmtContext> VisitResumeStmt([NotNull] VBAParser.ResumeStmtContext context)
            {
                if (_qualifiedSelection.Selection.Contains(context))
                {
                    InvalidContexts.Add(Tuple.Create(context as ParserRuleContext, "Extract method cannot extract methods that contains a Resume statement"));
                    return null;
                }

                return base.VisitResumeStmt(context);
            }

            public override IEnumerable<VBAParser.BlockStmtContext> VisitReturnStmt([NotNull] VBAParser.ReturnStmtContext context)
            {
                if (_qualifiedSelection.Selection.Contains(context))
                {
                    InvalidContexts.Add(Tuple.Create(context as ParserRuleContext, "Extract method cannot extract methods that contains a Return statement"));
                    return null;
                }

                return VisitReturnStmt(context);
            }
        }
    }
}
