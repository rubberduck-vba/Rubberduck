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
        private readonly IIndenterSettings _indenterSettings;
    
        public CodeMetricsAnalyst(IIndenterSettings indenterSettings)
        {
            _indenterSettings = indenterSettings;
        }

        public IEnumerable<ModuleMetricsResult> ModuleMetrics(RubberduckParserState state)
        {
            if (state == null || !state.AllUserDeclarations.Any())
            {
                // must not return Enumerable.Empty
                yield break;
            }

            var trees = state.ParseTrees;

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
            var cmListener = new CodeMetricsListener(declarationFinder, _indenterSettings);
            ParseTreeWalker.Default.Walk(cmListener, moduleTree);
            return cmListener.GetMetricsResult(qmn);
        }


        private class CodeMetricsListener : VBAParserBaseListener
        {
            private readonly DeclarationFinder _finder;
            private readonly IIndenterSettings _indenterSettings;

            private Declaration _currentMember;
            private List<CodeMetricsResult> _results = new List<CodeMetricsResult>();
            private List<CodeMetricsResult> _moduleResults = new List<CodeMetricsResult>();

            private List<MemberMetricsResult> _memberResults = new List<MemberMetricsResult>();

            public CodeMetricsListener(DeclarationFinder finder, IIndenterSettings indenterSettings)
            {
                _finder = finder;
                _indenterSettings = indenterSettings;
            }

            public override void EnterEndOfLine([NotNull] VBAParser.EndOfLineContext context)
            {
                int followingIndentationLevel = 0;
                // we have a proper newline
                if (context.NEWLINE() != null)
                {
                    // the last whitespace, which is the one in front of the next line's contents
                    var followingWhitespace = context.whiteSpace().LastOrDefault();
                    followingIndentationLevel = IndentationLevelFromWhitespace(followingWhitespace);
                }
                (_currentMember == null ? _moduleResults : _results).Add(new CodeMetricsResult(1, 0, followingIndentationLevel));
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

            public override void EnterBlockStmt([NotNull] VBAParser.BlockStmtContext context)
            {
                // there is a whitespace context here after the option of a statementLabel.
                // we need to account for that
                _results.Add(new CodeMetricsResult(0, 0, IndentationLevelFromWhitespace(context.whiteSpace())));
            }
            
            private int IndentationLevelFromWhitespace(VBAParser.WhiteSpaceContext wsContext)
            {
                if (wsContext == null) return 0;
                // the only thing that contains underscores is the line-continuation at this point
                var lineContinuation = wsContext.children.LastOrDefault((tree) => tree.GetText().Contains("_"));
                var index = lineContinuation != null ? wsContext.children.IndexOf(lineContinuation) : 0;
                return (wsContext?.ChildCount ?? 0 - index) / _indenterSettings.IndentSpaces;
            }

            private void ExitMeasurableMember()
            {
                _memberResults.Add(new MemberMetricsResult(_currentMember, _results));
                _results = new List<CodeMetricsResult>(); // reinitialize to drop results
                _currentMember = null;
            }

            internal ModuleMetricsResult GetMetricsResult(QualifiedModuleName qmn)
            {
                return new ModuleMetricsResult(qmn, _memberResults, _moduleResults);
            }
        }
    }
}
