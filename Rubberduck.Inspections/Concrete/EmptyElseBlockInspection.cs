using Antlr4.Runtime.Misc;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Inspections.Concrete
{
    internal class EmptyElseBlockInspection : ParseTreeInspectionBase
    {
        public EmptyElseBlockInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.DoNotShow) { }

        public override CodeInspectionType InspectionType => CodeInspectionType.MaintainabilityAndReadabilityIssues;

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return Listener.Contexts
                .Where(result => !IsIgnoringInspectionResultFor(result.ModuleName, result.Context.Start.Line))
                .Select(result => new QualifiedContextInspectionResult(this,
                                                        InspectionsUI.EmptyElseBlockInspectionResultFormat,
                                                        result));
        }

        public override IInspectionListener Listener { get; } 
            = new EmptyElseBlockListener();
        
        public class EmptyElseBlockListener : EmptyBlockInspectionListenerBase
        {
            public override void EnterElseBlock([NotNull] VBAParser.ElseBlockContext context)
            {
                InspectBlockForExecutableStatements(context.block(), context);
            }
        }
    }
}
