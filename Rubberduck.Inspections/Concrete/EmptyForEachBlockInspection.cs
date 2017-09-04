using Antlr4.Runtime.Misc;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    internal class EmptyForEachBlockInspection : EmptyBlockInspectionBase<EmptyDoWhileBlockInspection>
    {
        public EmptyForEachBlockInspection(RubberduckParserState state)
            : base(state, InspectionsUI.EmptyForEachBlockInspectionFormat) { }

        public override IInspectionListener Listener { get; } =
            new EmptyForEachBlockListener();

        public class EmptyForEachBlockListener : EmptyBlockListenerBase
        {
            public override void EnterForEachStmt([NotNull] VBAParser.ForEachStmtContext context)
            {
                InspectBlockForExecutableStatements(context.block(), context);
            }
        }
    }
}