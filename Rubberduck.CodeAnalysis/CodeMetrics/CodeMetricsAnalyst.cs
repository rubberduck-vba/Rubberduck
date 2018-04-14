using Antlr4.Runtime.Misc;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace Rubberduck.CodeAnalysis.CodeMetrics
{
    public class CodeMetricsAnalyst : ICodeMetricsAnalyst
    {

        private readonly CodeMetric[] metrics;

        public CodeMetricsAnalyst(IEnumerable<CodeMetric> supportedMetrics)
        {
            metrics = supportedMetrics.ToArray();
        }

        public IEnumerable<ICodeMetricResult> GetMetrics(RubberduckParserState state)
        {
            if (state == null || !state.AllUserDeclarations.Any())
            {
                // can not explicitly return Enumerable.Empty, this is equivalent
                yield break;
            }

            var trees = state.ParseTrees;

            foreach (var result in trees.SelectMany(moduleTree => TraverseModuleTree(moduleTree.Value, state.DeclarationFinder, moduleTree.Key)))
            {
                yield return result;
            }
        }


        public IEnumerable<ICodeMetricResult> ModuleResults(QualifiedModuleName moduleName, RubberduckParserState state)
        {
            return TraverseModuleTree(state.GetParseTree(moduleName), state.DeclarationFinder, moduleName);
        }

        private IEnumerable<ICodeMetricResult> TraverseModuleTree(IParseTree parseTree, DeclarationFinder declarationFinder, QualifiedModuleName moduleName)
        {
            var listeners = metrics.Select(metric => metric.TreeListener).ToList();
            foreach (var l in listeners)
            {
                l.Reset();
                l.InjectContext((declarationFinder, moduleName));
            }
            var combinedMetricListener = new CombinedParseTreeListener(listeners);

            ParseTreeWalker.Default.Walk(combinedMetricListener, parseTree);
            return listeners.SelectMany(l => l.Results());
        }

        private ModuleMetricsResult GetModuleResult(QualifiedModuleName qmn, IParseTree moduleTree, DeclarationFinder declarationFinder)
        {
            // FIXME rewrite as visitor, see discussion on pulls#3522
            // That should make subtrees easier and allow us to expand metrics
            var cmListener = new CodeMetricsListener(declarationFinder);
            ParseTreeWalker.Default.Walk(cmListener, moduleTree);
            return cmListener.GetMetricsResult(qmn);
        }

        private class CodeMetricsListener : VBAParserBaseListener
        {
            private readonly DeclarationFinder _finder;

            private Declaration _currentMember;
            private int _currentNestingLevel = 0;
            private int _currentMaxNesting = 0;
            private List<ICodeMetricsResult> _results = new List<ICodeMetricsResult>();
            private List<ICodeMetricsResult> _moduleResults = new List<ICodeMetricsResult>();

            private List<IMemberMetricsResult> _memberResults = new List<IMemberMetricsResult>();

            public CodeMetricsListener(DeclarationFinder finder)
            {
                _finder = finder;
            }
            public override void EnterBlock([NotNull] VBAParser.BlockContext context)
            {
                _currentNestingLevel++;
                if (_currentNestingLevel > _currentMaxNesting)
                {
                    _currentMaxNesting = _currentNestingLevel;
                }
            }

            public override void ExitBlock([NotNull] VBAParser.BlockContext context)
            {
                _currentNestingLevel--;
            }

            public override void EnterEndOfLine([NotNull] VBAParser.EndOfLineContext context)
            {
                (_currentMember == null ? _moduleResults : _results).Add(new CodeMetricsResult(1, 0, 0));
            }

            public override void EnterIfStmt([NotNull] VBAParser.IfStmtContext context)
            {
                _results.Add(new CodeMetricsResult(0, 1, 0));
            }

            public override void EnterElseIfBlock([NotNull] VBAParser.ElseIfBlockContext context)
            {
                _results.Add(new CodeMetricsResult(0, 1, 0));
            }

            // notably: NO additional complexity for an Else-Block

            public override void EnterForEachStmt([NotNull] VBAParser.ForEachStmtContext context)
            {
                _results.Add(new CodeMetricsResult(0, 1, 0));
            }

            public override void EnterForNextStmt([NotNull] VBAParser.ForNextStmtContext context)
            {
                _results.Add(new CodeMetricsResult(0, 1, 0));
            }

            public override void EnterCaseClause([NotNull] VBAParser.CaseClauseContext context)
            {
                _results.Add(new CodeMetricsResult(0, 1, 0));
            }

            public override void EnterSubStmt([NotNull] VBAParser.SubStmtContext context)
            {
                _results.Add(new CodeMetricsResult(0, 1, 0));
                _currentMember = _finder.UserDeclarations(DeclarationType.Procedure).Where(d => d.Context == context).First();
            }

            public override void ExitSubStmt([NotNull] VBAParser.SubStmtContext context)
            {
                ExitMeasurableMember();
            }

            public override void EnterFunctionStmt([NotNull] VBAParser.FunctionStmtContext context)
            {
                _results.Add(new CodeMetricsResult(0, 1, 0));
                _currentMember = _finder.UserDeclarations(DeclarationType.Function).Where(d => d.Context == context).First();
            }

            public override void ExitFunctionStmt([NotNull] VBAParser.FunctionStmtContext context)
            {
                ExitMeasurableMember();
            }

            public override void EnterPropertyGetStmt([NotNull] VBAParser.PropertyGetStmtContext context)
            {
                _results.Add(new CodeMetricsResult(0, 1, 0));
                _currentMember = _finder.UserDeclarations(DeclarationType.PropertyGet).Where(d => d.Context == context).First();
            }

            public override void ExitPropertyGetStmt([NotNull] VBAParser.PropertyGetStmtContext context)
            {
                ExitMeasurableMember();
            }

            public override void EnterPropertyLetStmt([NotNull] VBAParser.PropertyLetStmtContext context)
            {
                _results.Add(new CodeMetricsResult(0, 1, 0));
                _currentMember = _finder.UserDeclarations(DeclarationType.PropertyLet).Where(d => d.Context == context).First();
            }

            public override void ExitPropertyLetStmt([NotNull] VBAParser.PropertyLetStmtContext context)
            {
                ExitMeasurableMember();
            }

            public override void EnterPropertySetStmt([NotNull] VBAParser.PropertySetStmtContext context)
            {
                _results.Add(new CodeMetricsResult(0, 1, 0));
                _currentMember = _finder.UserDeclarations(DeclarationType.PropertySet).Where(d => d.Context == context).First();
            }

            public override void ExitPropertySetStmt([NotNull] VBAParser.PropertySetStmtContext context)
            { 
                ExitMeasurableMember();
            }
            
            private void ExitMeasurableMember()
            {
                Debug.Assert(_currentNestingLevel == 0, "Unexpected Nesting Level when exiting Measurable Member");
                _results.Add(new CodeMetricsResult(0, 0, _currentMaxNesting));
                _memberResults.Add(new MemberMetricsResult(_currentMember, _results));
                // reset state
                _results = new List<ICodeMetricsResult>(); 
                _currentMaxNesting = 0;
                _currentMember = null;
            }

            internal ModuleMetricsResult GetMetricsResult(QualifiedModuleName qmn)
            {
                return new ModuleMetricsResult(qmn, _memberResults, _moduleResults);
            }
        }
    }
}
