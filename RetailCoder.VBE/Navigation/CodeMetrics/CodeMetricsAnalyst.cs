using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.VBA;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;
using Antlr4.Runtime.Misc;
using Rubberduck.Parsing.Symbols;
using Rubberduck.SmartIndenter;

namespace Rubberduck.Navigation.CodeMetrics
{
    public class CodeMetricsAnalyst : ICodeMetricsAnalyst
    {
        private readonly IIndenterSettings indenterSettings;
    
        public CodeMetricsAnalyst(IIndenterSettings indenterSettings)
        {
            this.indenterSettings = indenterSettings;
        }

        public IEnumerable<ModuleMetricsResult> ModuleMetrics(RubberduckParserState state)
        {
            if (state == null || !state.AllUserDeclarations.Any())
            {
                yield break;
            }

            var trees = state.ParseTrees;
            var results = new List<CodeMetricsResult>();

            foreach (var moduleTree in trees)
            {
                yield return GetModuleResult(moduleTree.Key, moduleTree.Value, state.DeclarationFinder);
            };
        }

        public ModuleMetricsResult GetModuleResult(RubberduckParserState state, QualifiedModuleName qmn)
        {
            return GetModuleResult(qmn, state.GetParseTree(qmn), state.DeclarationFinder);
        }

        private ModuleMetricsResult GetModuleResult(QualifiedModuleName qmn, IParseTree moduleTree, DeclarationFinder declarationFinder)
        {
            // Consider rewrite as visitor? That should make subtrees easier and allow us to expand metrics
            var cmListener = new CodeMetricsListener(declarationFinder, indenterSettings);
            ParseTreeWalker.Default.Walk(cmListener, moduleTree);
            return cmListener.GetMetricsResult(qmn);
        }


        private class CodeMetricsListener : VBAParserBaseListener
        {
            private readonly DeclarationFinder _finder;
            private readonly IIndenterSettings _indenterSettings;

            private Declaration currentMember;
            private List<CodeMetricsResult> results = new List<CodeMetricsResult>();
            private List<CodeMetricsResult> moduleResults = new List<CodeMetricsResult>();

            private List<MemberMetricsResult> memberResults = new List<MemberMetricsResult>();

            public CodeMetricsListener(DeclarationFinder finder, IIndenterSettings indenterSettings)
            {
                _finder = finder;
                _indenterSettings = indenterSettings;
            }

            public override void EnterEndOfLine([NotNull] VBAParser.EndOfLineContext context)
            {
                (currentMember == null ? moduleResults : results).Add(new CodeMetricsResult(1, 0, 0));
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

                // if First borks, we got a bigger problems
                currentMember = _finder.DeclarationsWithType(DeclarationType.Procedure).Where(d => d.Context == context).First();
            }

            public override void ExitSubStmt([NotNull] VBAParser.SubStmtContext context)
            {
                // well, we're done here
                ExitMeasurableMember();
            }

            public override void EnterFunctionStmt([NotNull] VBAParser.FunctionStmtContext context)
            {
                // this is the default path through the function
                results.Add(new CodeMetricsResult(0, 1, 0));

                // if First borks, we got bigger problems
                currentMember = _finder.DeclarationsWithType(DeclarationType.Function).Where(d => d.Context == context).First();
            }

            public override void ExitFunctionStmt([NotNull] VBAParser.FunctionStmtContext context)
            {
                // well, we're done here
                ExitMeasurableMember();
            }

            public override void EnterPropertyGetStmt([NotNull] VBAParser.PropertyGetStmtContext context)
            {
                results.Add(new CodeMetricsResult(0, 1, 0));

                currentMember = _finder.DeclarationsWithType(DeclarationType.PropertyGet).Where(d => d.Context == context).First();
            }

            public override void ExitPropertyGetStmt([NotNull] VBAParser.PropertyGetStmtContext context)
            {
                // well, we're done here
                ExitMeasurableMember();
            }

            public override void EnterPropertyLetStmt([NotNull] VBAParser.PropertyLetStmtContext context)
            {
                results.Add(new CodeMetricsResult(0, 1, 0));

                currentMember = _finder.DeclarationsWithType(DeclarationType.PropertyLet).Where(d => d.Context == context).First();
            }

            public override void ExitPropertyLetStmt([NotNull] VBAParser.PropertyLetStmtContext context)
            {
                // well, we're done here
                ExitMeasurableMember();
            }

            public override void EnterPropertySetStmt([NotNull] VBAParser.PropertySetStmtContext context)
            {
                results.Add(new CodeMetricsResult(0, 1, 0));

                currentMember = _finder.DeclarationsWithType(DeclarationType.PropertySet).Where(d => d.Context == context).First();
            }

            public override void ExitPropertySetStmt([NotNull] VBAParser.PropertySetStmtContext context)
            {
                // well, we're done here
                ExitMeasurableMember();
            }

            public override void EnterBlockStmt([NotNull] VBAParser.BlockStmtContext context)
            {
                var ws = context.whiteSpace();
                // FIXME only take the last contiguous non-linebreak into account
                results.Add(new CodeMetricsResult(0, 0, ws.ChildCount / _indenterSettings.IndentSpaces));
            }

            private void ExitMeasurableMember()
            {
                memberResults.Add(new MemberMetricsResult(currentMember, results));
                results = new List<CodeMetricsResult>(); // reinitialize to drop results
                currentMember = null;
            }

            internal ModuleMetricsResult GetMetricsResult(QualifiedModuleName qmn)
            {
                return new ModuleMetricsResult(qmn, memberResults, moduleResults);
            }
        }
    }
}
