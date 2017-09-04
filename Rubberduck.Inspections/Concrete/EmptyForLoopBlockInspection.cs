using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime.Misc;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;
using Rubberduck.Inspections.Concrete;

namespace RubberduckTests.Inspections
{
    internal class EmptyForLoopBlockInspection : ParseTreeInspectionBase
    {
        public EmptyForLoopBlockInspection(RubberduckParserState state)
            : base(state) { }

        public override Type Type => typeof(EmptyForLoopBlockInspection);

        public override CodeInspectionType InspectionType => CodeInspectionType.CodeQualityIssues;

        public override IInspectionListener Listener { get; } =
            new EmptyForloopBlockListener();

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            //TODO: create InspectionUI resource
            return Listener.Contexts
                .Where(result => !IsIgnoringInspectionResultFor(result.ModuleName, result.Context.Start.Line))
                .Select(result => new QualifiedContextInspectionResult(this,
                                                       //InspectionsUI.EmptyIfBlockInspectionResultFormat,
                                                       "For loop contains no executable statements",
                                                       result));
        }

        public class EmptyForloopBlockListener : EmptyBlockListenerBase
        {
            public override void EnterForNextStmt([NotNull] VBAParser.ForNextStmtContext context)
            {
                InspectBlockForExecutableStatements(context.block(), context);
            }
        }
    }
}
