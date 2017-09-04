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
    internal class EmptyWhileWendBlockInspection : ParseTreeInspectionBase
    {
        public EmptyWhileWendBlockInspection(RubberduckParserState state)
            : base(state) { }

        public override Type Type => typeof(EmptyWhileWendBlockInspection);

        public override CodeInspectionType InspectionType => CodeInspectionType.CodeQualityIssues;

        public override IInspectionListener Listener { get; } =
            new EmptyWhileWendBlockListener();

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            //TODO: create InspectionUI resource
            return Listener.Contexts
                .Where(result => !IsIgnoringInspectionResultFor(result.ModuleName, result.Context.Start.Line))
                .Select(result => new QualifiedContextInspectionResult(this,
                                                       //InspectionsUI.EmptyIfBlockInspectionResultFormat,
                                                       "While Wend loop contains no executable statements",
                                                       result));
        }

        public class EmptyWhileWendBlockListener : EmptyBlockListenerBase
        {
            public override void EnterWhileWendStmt([NotNull] VBAParser.WhileWendStmtContext context)
            {
                InspectBlockForExecutableStatements(context.block(), context);
            }
        }
    }
}
