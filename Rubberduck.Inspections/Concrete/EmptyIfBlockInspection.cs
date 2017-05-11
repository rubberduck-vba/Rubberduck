using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Concrete
{
    internal class EmptyIfBlockInspection : InspectionBase, IParseTreeInspection
    {
        public EmptyIfBlockInspection(RubberduckParserState state)
            : base(state) { }

        public override CodeInspectionType InspectionType => CodeInspectionType.CodeQualityIssues;

        public IInspectionListener Listener { get; } =
            new EmptyStringLiteralListener();

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            return Listener.Contexts
                .Where(result => !IsIgnoringInspectionResultFor(result.ModuleName, result.Context.Start.Line))
                .Select(result => new QualifiedContextInspectionResult(this,
                                                       InspectionsUI.EmptyIfBlockInspectionResultFormat,
                                                       State,
                                                       result));
        }

        public class EmptyStringLiteralListener : VBAParserBaseListener, IInspectionListener
        {
            private readonly List<QualifiedContext<ParserRuleContext>> _contexts = new List<QualifiedContext<ParserRuleContext>>();
            public IReadOnlyList<QualifiedContext<ParserRuleContext>> Contexts => _contexts;

            public QualifiedModuleName CurrentModuleName { get; set; }

            public void ClearContexts()
            {
                _contexts.Clear();
            }

            public override void EnterIfStmt([NotNull] VBAParser.IfStmtContext context)
            {
                var block = context.block();
                if (block == null || block.children == null || !ContainsExecutableStatements(block))
                {
                    _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
                }
            }

            public override void EnterElseIfBlock([NotNull] VBAParser.ElseIfBlockContext context)
            {
                var block = context.block();
                if (block == null || block.children == null || !ContainsExecutableStatements(block))
                {
                    _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
                }
            }

            public override void EnterSingleLineIfStmt([NotNull] VBAParser.SingleLineIfStmtContext context)
            {
                if (context.ifWithEmptyThen() != null)
                {
                    _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context.ifWithEmptyThen()));
                }
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
                            continue;   //We have a lone line lable, which is not executable.
                        }

                        Debug.Assert(mainBlockStmt.ChildCount == 1);

                        // exclude variables and consts because they are not executable statements
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
