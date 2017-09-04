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

namespace Rubberduck.Inspections.Concrete
{
    internal class EmptyForEachBlockInspection : ParseTreeInspectionBase
    {
        public EmptyForEachBlockInspection(RubberduckParserState state)
            : base(state) { }

        public override Type Type => typeof(EmptyForEachBlockInspection);

        public override CodeInspectionType InspectionType => CodeInspectionType.CodeQualityIssues;

        public override IInspectionListener Listener { get; } =
            new EmptyCaseBlockListener();

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            //TODO: create InspectionUI resource
            return Listener.Contexts
                .Where(result => !IsIgnoringInspectionResultFor(result.ModuleName, result.Context.Start.Line))
                .Select(result => new QualifiedContextInspectionResult(this,
                                                       //InspectionsUI.EmptyIfBlockInspectionResultFormat,
                                                       "For Each loop contains no executable statements",
                                                       result));
        }

        public class EmptyCaseBlockListener : EmptyBlockListenerBase
        {

            public override void EnterForEachStmt([NotNull] VBAParser.ForEachStmtContext context)
            {
                InspectBlockForExecutableStatements(context.block(), context);
            }
        }
    }
}
