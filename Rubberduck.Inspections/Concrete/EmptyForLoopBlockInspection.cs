using Antlr4.Runtime.Misc;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    internal class EmptyForLoopBlockInspection : EmptyBlockInspectionBase<EmptyForLoopBlockInspection>
    {
        public EmptyForLoopBlockInspection(RubberduckParserState state)
            : base(state, InspectionsUI.EmptyForLoopBlockInspectionFormat) { }

        public override IInspectionListener Listener { get; } =
            new EmptyForloopBlockListener();

        public class EmptyForloopBlockListener : EmptyBlockListenerBase
        {
            public override void EnterForNextStmt([NotNull] VBAParser.ForNextStmtContext context)
            {
                InspectBlockForExecutableStatements(context.block(), context);
            }
        }
    }
}
