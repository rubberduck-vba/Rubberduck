using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;
using Antlr4.Runtime.Misc;

namespace Rubberduck.Inspections.Concrete
{
    internal class EmptyElseBlockInspection : ParseTreeInspectionBase
    {
        public EmptyElseBlockInspection(RubberduckParserState state) : base(state, CodeInspectionSeverity.Suggestion) { }

        public override Type Type => typeof(EmptyElseBlockInspection);

        public override CodeInspectionType InspectionType => CodeInspectionType.CodeQualityIssues;

        public override IInspectionListener Listener { get; } = new EmptyElseBlockListener();
        
        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            return Listener.Contexts
                .Where(result => !IsIgnoringInspectionResultFor(result.ModuleName, result.Context.Start.Line))
                .Select(result => new QualifiedContextInspectionResult(this,
                                                        InspectionsUI.EmptyElseBlockInspectionResultFormat,
                                                        result));
        }

        public class EmptyElseBlockListener : EmptyBlockListenerBase
        {
            public override void EnterElseBlock([NotNull] VBAParser.ElseBlockContext context)
            {
                InspectBlockForExecutableStatements(context.block(), context);
            }
        }
    }
}