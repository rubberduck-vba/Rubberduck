using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.Refactorings.ExtractMethod
{
    public class ExtractMethodSelectionValidation : IExtractMethodSelectionValidation
    {
        private readonly IProjectsProvider _projectsProvider;
        private readonly IEnumerable<Declaration> _declarations;
        private List<Tuple<ParserRuleContext, string>> _invalidContexts = new List<Tuple<ParserRuleContext, string>>();
        private List<VBABaseParserRuleContext> _finalResults = new List<VBABaseParserRuleContext>();

        public ExtractMethodSelectionValidation(IEnumerable<Declaration> declarations, IProjectsProvider projectsProvider)
        {
            _declarations = declarations;
            _projectsProvider = projectsProvider;
        }

        public IEnumerable<Tuple<ParserRuleContext, string>> InvalidContexts => _invalidContexts;

        public IEnumerable<VBABaseParserRuleContext> SelectedContexts => _finalResults;

        public static readonly DeclarationType[] ProcedureTypes =
        {
            DeclarationType.Procedure,
            DeclarationType.Function,
            DeclarationType.PropertyGet,
            DeclarationType.PropertyLet,
            DeclarationType.PropertySet
        };
        private Declaration ProcDeclaration(QualifiedSelection qualifiedSelection, int lineOfInterest)
        {
            var selection = qualifiedSelection.Selection;
            var procedures = _declarations.Where(d => d.ComponentName == qualifiedSelection.QualifiedName.ComponentName && d.IsUserDefined && (ProcedureTypes.Contains(d.DeclarationType)));
            var declarations = procedures as IList<Declaration> ?? procedures.ToList();
            Declaration ProcOfLine(int sl) => declarations.FirstOrDefault(d => d.Context.Start.Line < sl && d.Context.Stop.EndLine() > sl);
            return ProcOfLine(lineOfInterest);
        }
            
        public bool IsSelectionValid(QualifiedSelection qualifiedSelection)
        {
            var selection = qualifiedSelection.Selection;

            var startLine = selection.StartLine;
            var endLine = selection.EndLine;

            var procEnd = ProcDeclaration(qualifiedSelection, endLine);
            if (procEnd == null)
            {
                _invalidContexts.Add(Tuple.Create(null as ParserRuleContext, RefactoringsUI.ExtractMethod_InvalidMessageSelectionEndsOutsideProcedure));
                return false;
            }

            var procStart = ProcDeclaration(qualifiedSelection, startLine);
            if (procStart == null)
            {
                _invalidContexts.Add(Tuple.Create(null as ParserRuleContext, RefactoringsUI.ExtractMethod_InvalidMessageSelectionStartsOutsideProcedure));
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
                    _invalidContexts.Add(Tuple.Create(procStartContext, RefactoringsUI.ExtractMethod_InvalidMessageSelectionNotInRecognisedProcedure));
                    return false;
            }

            if (!(procEnd.QualifiedSelection.Equals(procStart.QualifiedSelection)
                && (procEndOfSignature.Start.Line < selection.StartLine
                || procEndOfSignature.Start.Line == selection.StartLine && procEndOfSignature.Start.Column < selection.StartColumn)
                ))
            {
                _invalidContexts.Add(Tuple.Create(null as ParserRuleContext, RefactoringsUI.ExtractMethod_InvalidMessageSelectionMoreThanSingleProcedure));
                return false;
            }

            /* At this point, we know the selection is within a single procedure. We need to validate that the user's
             * selection in fact contain only BlockStmt and not other stuff that might not be so extractable.
             */
            var visitor = new ExtractValidatorVisitor(qualifiedSelection, _invalidContexts);
            var results = visitor.Visit(procStartContext);
            var endOfStatementContexts = visitor.EndOfStatementContexts;
            endOfStatementContexts.Add(procEndOfSignature);
            _invalidContexts = visitor.InvalidContexts;

            if (!_invalidContexts.Any())
            {
                using (var component = _projectsProvider.Component(qualifiedSelection.QualifiedName))
                {
                    if (component.CodeModule.ContainsCompilationDirectives(selection))
                    {
                        ContainsCompilerDirectives = true;
                    }
                }
                // We've proved that there are no invalid statements contained in the selection. However, we need to analyze
                // the statements to ensure they are not partial selections.

                // The visitor will not return the results in a sorted manner, so we need to arrange the contexts in the same order.
                var blockStmtContexts = results as IList<VBAParser.BlockStmtContext> ?? results.ToList();
                var sorted = blockStmtContexts.OrderBy(context => context.Start.StartIndex);
                ContextIsContainedOnce(sorted, endOfStatementContexts, ref _finalResults, qualifiedSelection);
                return blockStmtContexts.Any() && !_invalidContexts.Any() && _finalResults.Any();
            }
            return false;
        }

        public bool ContainsCompilerDirectives { get; set; }

        /// <summary>
        /// The function ensure that we return only top-level BlockStmtContexts and EndOfStatements that
        /// exist within a user's selection, excluding any nested contexts which are also "selected" and 
        /// thus ensure that we build an unique list of contexts that corresponds to the user's selection.
        /// The function will also validate that there are no overlapping selections which could be invalid.
        /// </summary>
        /// <param name="sortedResults">The BlockStmtContexts to test</param>
        /// <param name="endOfStatementContexts">The endOfStatementContexts to test</param>
        /// <param name="aggregate">The list of contexts we already added to verify we are not adding one of its children or itself more than once</param>
        /// <param name="qualifiedSelection"></param>
        private void ContextIsContainedOnce(IEnumerable<VBAParser.BlockStmtContext> sortedResults, IEnumerable<VBAParser.EndOfStatementContext> endOfStatementContexts, ref List<VBABaseParserRuleContext> aggregate, QualifiedSelection qualifiedSelection)
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
                        _invalidContexts.Add(Tuple.Create(context as ParserRuleContext, RefactoringsUI.ExtractMethod_InvalidMessageSelectionNotSetOfCompleteStatements));
                    }
                }
            }
            foreach (var context in endOfStatementContexts.OrderBy(c => c.Start.StartIndex))
            {
                if (qualifiedSelection.Selection.Contains(context))
                {
                    if (!aggregate.Any(otherContext => otherContext.GetSelection().Contains(context) && context != otherContext))
                    {
                        aggregate.Add(context);
                    }
                }
            }
            //Final sort in case any blank or comment only lines before the first block statements
            aggregate.Sort((a, b) => a.Start.StartIndex.CompareTo(b.Start.StartIndex));
        }

        private class ExtractValidatorVisitor : VBAParserBaseVisitor<IEnumerable<VBAParser.BlockStmtContext>>
        {
            private readonly QualifiedSelection _qualifiedSelection;

            public ExtractValidatorVisitor(QualifiedSelection qualifiedSelection, List<Tuple<ParserRuleContext, string>> invalidContexts)
            {
                _qualifiedSelection = qualifiedSelection;
                InvalidContexts = invalidContexts;
                EndOfStatementContexts = new List<VBAParser.EndOfStatementContext>();
            }

            public List<Tuple<ParserRuleContext, string>> InvalidContexts { get; }
            public List<VBAParser.EndOfStatementContext> EndOfStatementContexts { get; }

            protected override IEnumerable<VBAParser.BlockStmtContext> DefaultResult => new List<VBAParser.BlockStmtContext>();
            public override IEnumerable<VBAParser.BlockStmtContext> VisitEndOfStatement([NotNull] VBAParser.EndOfStatementContext context)
            {
                if (_qualifiedSelection.Selection.Contains(context))
                {
                    EndOfStatementContexts.Add(context);
                }
                return base.VisitEndOfStatement(context);
            }
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
                    InvalidContexts.Add(Tuple.Create(context as ParserRuleContext, string.Format(RefactoringsUI.ExtractMethod_InvalidMessageSelectionHasUnsupportedStatement, "Error")));
                    return null;
                }

                return base.VisitErrorStmt(context);
            }

            public override IEnumerable<VBAParser.BlockStmtContext> VisitEndStmt([NotNull] VBAParser.EndStmtContext context)
            {
                if (_qualifiedSelection.Selection.Contains(context))
                {
                    InvalidContexts.Add(Tuple.Create(context as ParserRuleContext, string.Format(RefactoringsUI.ExtractMethod_InvalidMessageSelectionHasUnsupportedStatement, "End")));
                    return null;
                }

                return base.VisitEndStmt(context);
            }

            public override IEnumerable<VBAParser.BlockStmtContext> VisitExitStmt([NotNull] VBAParser.ExitStmtContext context)
            {
                if (_qualifiedSelection.Selection.Contains(context))
                {
                    InvalidContexts.Add(Tuple.Create(context as ParserRuleContext, string.Format(RefactoringsUI.ExtractMethod_InvalidMessageSelectionHasUnsupportedStatement, "Exit")));
                    return null;
                }

                return base.VisitExitStmt(context);
            }

            public override IEnumerable<VBAParser.BlockStmtContext> VisitGoSubStmt([NotNull] VBAParser.GoSubStmtContext context)
            {
                if (_qualifiedSelection.Selection.Contains(context))
                {
                    InvalidContexts.Add(Tuple.Create(context as ParserRuleContext, string.Format(RefactoringsUI.ExtractMethod_InvalidMessageSelectionHasUnsupportedStatement, "GoSub")));
                    return null;
                }

                return base.VisitGoSubStmt(context);
            }

            public override IEnumerable<VBAParser.BlockStmtContext> VisitGoToStmt([NotNull] VBAParser.GoToStmtContext context)
            {
                if (_qualifiedSelection.Selection.Contains(context))
                {
                    InvalidContexts.Add(Tuple.Create(context as ParserRuleContext, string.Format(RefactoringsUI.ExtractMethod_InvalidMessageSelectionHasUnsupportedStatement, "GoTo")));
                    return null;
                }

                return base.VisitGoToStmt(context);
            }

            public override IEnumerable<VBAParser.BlockStmtContext> VisitOnErrorStmt([NotNull] VBAParser.OnErrorStmtContext context)
            {
                if (_qualifiedSelection.Selection.Contains(context))
                {
                    InvalidContexts.Add(Tuple.Create(context as ParserRuleContext, string.Format(RefactoringsUI.ExtractMethod_InvalidMessageSelectionHasUnsupportedStatement, "On Error")));
                    return null;
                }

                return base.VisitOnErrorStmt(context);
            }
            
            public override IEnumerable<VBAParser.BlockStmtContext> VisitIdentifierStatementLabel([NotNull] VBAParser.IdentifierStatementLabelContext context)
            {
                if (_qualifiedSelection.Selection.Contains(context))
                {
                    InvalidContexts.Add(Tuple.Create(context as ParserRuleContext, string.Format(RefactoringsUI.ExtractMethod_InvalidMessageSelectionHasUnsupportedStatement, "line label")));
                    return null;
                }

                return base.VisitIdentifierStatementLabel(context);
            }

            public override IEnumerable<VBAParser.BlockStmtContext> VisitCombinedLabels([NotNull] VBAParser.CombinedLabelsContext context)
            {
                if (_qualifiedSelection.Selection.Contains(context))
                {
                    InvalidContexts.Add(Tuple.Create(context as ParserRuleContext, string.Format(RefactoringsUI.ExtractMethod_InvalidMessageSelectionHasUnsupportedStatement, "line label")));
                    return base.VisitCombinedLabels(context);
                }

                return base.VisitCombinedLabels(context);
            }

            public override IEnumerable<VBAParser.BlockStmtContext> VisitStandaloneLineNumberLabel([NotNull] VBAParser.StandaloneLineNumberLabelContext context)
            {
                if (_qualifiedSelection.Selection.Contains(context))
                {
                    InvalidContexts.Add(Tuple.Create(context as ParserRuleContext, string.Format(RefactoringsUI.ExtractMethod_InvalidMessageSelectionHasUnsupportedStatement, "line label")));
                    return null;
                }

                return base.VisitStandaloneLineNumberLabel(context);
            }

            public override IEnumerable<VBAParser.BlockStmtContext> VisitOnGoSubStmt([NotNull] VBAParser.OnGoSubStmtContext context)
            {
                if (_qualifiedSelection.Selection.Contains(context))
                {
                    InvalidContexts.Add(Tuple.Create(context as ParserRuleContext, string.Format(RefactoringsUI.ExtractMethod_InvalidMessageSelectionHasUnsupportedStatement, "On...GoSub")));
                    return null;
                }

                return base.VisitOnGoSubStmt(context);
            }

            public override IEnumerable<VBAParser.BlockStmtContext> VisitOnGoToStmt([NotNull] VBAParser.OnGoToStmtContext context)
            {
                if (_qualifiedSelection.Selection.Contains(context))
                {
                    InvalidContexts.Add(Tuple.Create(context as ParserRuleContext, string.Format(RefactoringsUI.ExtractMethod_InvalidMessageSelectionHasUnsupportedStatement, "On...GoTo")));
                    return null;
                }

                return base.VisitOnGoToStmt(context);
            }

            public override IEnumerable<VBAParser.BlockStmtContext> VisitResumeStmt([NotNull] VBAParser.ResumeStmtContext context)
            {
                if (_qualifiedSelection.Selection.Contains(context))
                {
                    InvalidContexts.Add(Tuple.Create(context as ParserRuleContext, string.Format(RefactoringsUI.ExtractMethod_InvalidMessageSelectionHasUnsupportedStatement, "Resume")));
                    return null;
                }

                return base.VisitResumeStmt(context);
            }

            public override IEnumerable<VBAParser.BlockStmtContext> VisitReturnStmt([NotNull] VBAParser.ReturnStmtContext context)
            {
                if (_qualifiedSelection.Selection.Contains(context))
                {
                    InvalidContexts.Add(Tuple.Create(context as ParserRuleContext, string.Format(RefactoringsUI.ExtractMethod_InvalidMessageSelectionHasUnsupportedStatement, "Return")));
                    return null;
                }

                return base.VisitReturnStmt(context);
            }
        }
    }
}
