using Antlr4.Runtime.Misc;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    internal class EmptyDoWhileBlockInspection : EmptyBlockInspectionBase<EmptyDoWhileBlockInspection>
    {
        public EmptyDoWhileBlockInspection(RubberduckParserState state) 
            : base(state, InspectionsUI.EmptyDoWhileBlockInspectionFormat) { }

        public override IInspectionListener Listener { get; } =
            new EmptyDoWhileBlockListener();

        public class EmptyDoWhileBlockListener : EmptyBlockListenerBase
        {
            public override void EnterDoLoopStmt([NotNull] VBAParser.DoLoopStmtContext context)
            {
                InspectBlockForExecutableStatements(context.block(), context);
            }
        }
    }
}
