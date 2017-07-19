using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System.Diagnostics;
using Antlr4.Runtime.Misc;

namespace Rubberduck.Inspections.Concrete
{
    internal class EmptyElseBlockInspection : ParseTreeInspectionBase
    {
        public EmptyElseBlockInspection(RubberduckParserState state) : base(state, CodeInspectionSeverity.Suggestion) { }

        public override CodeInspectionType InspectionType => CodeInspectionType.CodeQualityIssues;

        public override IInspectionListener Listener { get; } = new EmptyElseBlockListener();
        
        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            return Listener.Contexts
                .Where(result => !IsIgnoringInspectionResultFor(result.ModuleName, result.Context.Start.Line))
                .Select(result => new QualifiedContextInspectionResult(this,
                                                        InspectionsUI.EmptyElseBlockInspectionResultFormat,
                                                        State,
                                                        result));
        }

        public class EmptyElseBlockListener : VBAParserBaseListener, IInspectionListener
        {
            public override void EnterElseBlock([NotNull] VBAParser.ElseBlockContext context)
            {
                var block = context.block();
                if (block == null || block.children == null || !ContainsExecutableStatements(block))
                {
                    _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
                }
                base.EnterElseBlock(context);
            }

            private readonly List<QualifiedContext<ParserRuleContext>> _contexts = new List<QualifiedContext<ParserRuleContext>>();
            public IReadOnlyList<QualifiedContext<ParserRuleContext>> Contexts => _contexts;

            public QualifiedModuleName CurrentModuleName { get; set; }

            public void ClearContexts()
            {
                _contexts.Clear();
            }

            private bool ContainsExecutableStatements(VBAParser.BlockContext block)
            {
                foreach (var child in block.children)
                {
                    if (child is VBAParser.BlockStmtContext)
                    {
                        var blockStmt = (VBAParser.BlockStmtContext)child;
                        var mainBlockStmt = blockStmt.mainBlockStmt();

                        if(mainBlockStmt == null)
                        {
                            continue; //Lone line label, which isn't executable.
                        }

                        Debug.Assert(mainBlockStmt.ChildCount == 1);

                        if (mainBlockStmt.GetChild(0) is VBAParser.VariableStmtContext ||
                            mainBlockStmt.GetChild(0) is VBAParser.ConstStmtContext)
                        {
                            continue;
                        }

                        return true;
                    }

                    if (child is VBAParser.RemCommentContext ||
                        child is VBAParser.CommentContext ||
                        child is VBAParser.CommentOrAnnotationContext ||
                        child is VBAParser.EndOfStatementContext)
                    {
                        continue;
                    }

                    return true;
                }

                return false;
            }
        }
    }
}