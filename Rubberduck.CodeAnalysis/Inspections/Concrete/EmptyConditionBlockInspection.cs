using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing;
using Rubberduck.Parsing.VBA;
using Rubberduck.Inspections.Results;

namespace Rubberduck.Inspections.Concrete
{

    internal class EmptyConditionBlockInspection : ParseTreeInspectionBase
    {
        public EmptyConditionBlockInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion) { }

        public override Type Type => typeof(EmptyConditionBlockInspection);

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            return Listener.Contexts
                .Where(result => !IsIgnoringInspectionResultFor(result.ModuleName, result.Context.Start.Line))
                .Select(result => new QualifiedContextInspectionResult(this,
                                                        InspectionsUI.EmptyConditionBlockInspectionsResultFormat,
                                                        result));
        }

        public override IInspectionListener Listener { get; } = 
            new EmptyConditionBlockListener();

        public class EmptyConditionBlockListener : EmptyBlockInspectionListenerBase
        {
            public override void EnterIfStmt([NotNull] VBAParser.IfStmtContext context)
            {
                InspectBlockForExecutableStatements(context.block(), context);
            }

            public override void EnterElseIfBlock([NotNull] VBAParser.ElseIfBlockContext context)
            {
                InspectBlockForExecutableStatements(context.block(), context);
            }

            public override void EnterSingleLineIfStmt([NotNull] VBAParser.SingleLineIfStmtContext context)
            {
                AddResult(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context.ifWithEmptyThen()));
            }

            public override void EnterElseBlock([NotNull] VBAParser.ElseBlockContext context)
            {
                InspectBlockForExecutableStatements(context.block(), context);
            }
        }
    }
}