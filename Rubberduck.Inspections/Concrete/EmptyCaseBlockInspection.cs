using Antlr4.Runtime.Misc;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    internal class EmptyCaseBlockInspection : EmptyBlockInspectionBase<EmptyCaseBlockInspection>
    {
        public EmptyCaseBlockInspection(RubberduckParserState state) 
            : base(state,  InspectionsUI.EmptyCaseBlockInspectionResultFormat) { }

        public override IInspectionListener Listener { get; } =
            new EmptyCaseBlockListener();

        public class EmptyCaseBlockListener : EmptyBlockListenerBase
        {
            public override void EnterCaseClause([NotNull] VBAParser.CaseClauseContext context)
            {
                InspectBlockForExecutableStatements(context.block(), context);
            }
        }
    }
}
