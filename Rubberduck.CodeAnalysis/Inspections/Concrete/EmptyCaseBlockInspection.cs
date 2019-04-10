using Antlr4.Runtime.Misc;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Common;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Resources.Experimentals;
using Rubberduck.Inspections.Inspections.Extensions;

namespace Rubberduck.Inspections.Concrete
{
    [Experimental(nameof(ExperimentalNames.EmptyBlockInspections))]
    internal class EmptyCaseBlockInspection : ParseTreeInspectionBase
    {
        public EmptyCaseBlockInspection(RubberduckParserState state)
            : base(state) { }

        public override IInspectionListener Listener { get; } =
            new EmptyCaseBlockListener();

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return Listener.Contexts
                .Where(result => !result.IsIgnoringInspectionResultFor(State.DeclarationFinder, AnnotationName))
                .Select(result => new QualifiedContextInspectionResult(this,
                                                        InspectionResults.EmptyCaseBlockInspection,
                                                        result));
        }

        public class EmptyCaseBlockListener : EmptyBlockInspectionListenerBase
        {
            public override void EnterCaseClause([NotNull] VBAParser.CaseClauseContext context)
            {
                InspectBlockForExecutableStatements(context.block(), context);
            }
        }
    }
}
