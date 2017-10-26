using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Rubberduck.Parsing.VBA;
using Antlr4.Runtime.Tree;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;
using Antlr4.Runtime.Misc;

namespace Rubberduck.Navigation.CodeMetrics
{
    public class ParseTreeMetricsAnalyst : ICodeMetricsAnalyst
    {
        public async Task<CodeMetricsResult> GetResult(RubberduckParserState state, CancellationToken token)
        {
            if (state == null || !state.AllUserDeclarations.Any())
            {
                return new CodeMetricsResult();
            }
            return await Task.Run(() =>
            {
                var trees = state.ParseTrees;
                var results = new List<CodeMetricsResult>();

                foreach (var moduleTree in trees)
                {
                    if (token.IsCancellationRequested)
                    {
                        return new CodeMetricsResult();
                    }
                    // FIXME rewrite as visitor. That should make subtrees easier and allow us to expand metrics
                    var cmListener = new CodeMetricsListener(moduleTree.Key);
                    ParseTreeWalker.Default.Walk(cmListener, moduleTree.Value);
                    results.Add(cmListener.GetMetricsResult());
                }
                return new CodeMetricsResult(0, 0, 0, results);
            });
        }

        private class CodeMetricsListener : VBAParserBaseListener
        {
            private QualifiedModuleName qmn;
            private List<CodeMetricsResult> results = new List<CodeMetricsResult>();

            public CodeMetricsListener(QualifiedModuleName qmn)
            {
                this.qmn = qmn;
            }

            public override void EnterEndOfLine([NotNull] VBAParser.EndOfLineContext context)
            {
                results.Add(new CodeMetricsResult(1, 0, 0));
            }

            public override void EnterIfStmt([NotNull] VBAParser.IfStmtContext context)
            {
                // one additional path beside the default
                results.Add(new CodeMetricsResult(0, 1, 0));
            }

            public override void EnterElseIfBlock([NotNull] VBAParser.ElseIfBlockContext context)
            {
                // one additonal path beside the default
                results.Add(new CodeMetricsResult(0, 1, 0));
            }

            public override void EnterForEachStmt([NotNull] VBAParser.ForEachStmtContext context)
            {
                // one additional path
                results.Add(new CodeMetricsResult(0, 1, 0));
            }

            public override void EnterForNextStmt([NotNull] VBAParser.ForNextStmtContext context)
            {
                results.Add(new CodeMetricsResult(0, 1, 0));
            }

            public override void EnterCaseClause([NotNull] VBAParser.CaseClauseContext context)
            {
                results.Add(new CodeMetricsResult(0, 1, 0));
            }

            public override void EnterSubStmt([NotNull] VBAParser.SubStmtContext context)
            {
                // this is the default path through the sub
                results.Add(new CodeMetricsResult(0, 1, 0));
            }

            public override void EnterFunctionStmt([NotNull] VBAParser.FunctionStmtContext context)
            {
                // this is the default path through the function
                results.Add(new CodeMetricsResult(0, 1, 0));
            }

            public override void EnterBlockStmt([NotNull] VBAParser.BlockStmtContext context)
            {
                var ws = context.whiteSpace();
                // FIXME divide by indent size and assume we're indented?
                // FIXME LINE_CONTINUATION might interfere here
                //results.Add(new CodeMetricsResult(0, 0, ws.ChildCount / 4));
            }

            // FIXME also check if we need to do something about `mandatoryLineContinuation`?

            internal CodeMetricsResult GetMetricsResult()
            {
                return new CodeMetricsResult(0, 0, 0, results);
            }
        }
    }
}
