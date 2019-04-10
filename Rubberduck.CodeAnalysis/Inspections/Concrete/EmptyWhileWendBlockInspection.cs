using Antlr4.Runtime.Misc;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Common;
using Rubberduck.Parsing.Grammar;
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
    internal class EmptyWhileWendBlockInspection : ParseTreeInspectionBase
    {
        public EmptyWhileWendBlockInspection(RubberduckParserState state)
            : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return Listener.Contexts
                .Where(result => !result.IsIgnoringInspectionResultFor(State.DeclarationFinder, AnnotationName))
                .Select(result => new QualifiedContextInspectionResult(this,
                                                        InspectionResults.EmptyWhileWendBlockInspection,
                                                        result));
        }

        public override IInspectionListener Listener { get; } =
            new EmptyWhileWendBlockListener();

        public class EmptyWhileWendBlockListener : EmptyBlockInspectionListenerBase
        {
            public override void EnterWhileWendStmt([NotNull] VBAParser.WhileWendStmtContext context)
            {
                InspectBlockForExecutableStatements(context.block(), context);
            }
        }
    }
}
