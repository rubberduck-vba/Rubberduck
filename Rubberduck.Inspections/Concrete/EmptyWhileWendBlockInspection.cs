using Antlr4.Runtime.Misc;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    internal class EmptyWhileWendBlockInspection : EmptyBlockInspectionBase<EmptyWhileWendBlockInspection>
    {
        public EmptyWhileWendBlockInspection(RubberduckParserState state)
            : base(state, "While Wend loop contains no executable statements") { }

        public override IInspectionListener Listener { get; } =
            new EmptyWhileWendBlockListener();

        public class EmptyWhileWendBlockListener : EmptyBlockListenerBase
        {
            public override void EnterWhileWendStmt([NotNull] VBAParser.WhileWendStmtContext context)
            {
                InspectBlockForExecutableStatements(context.block(), context);
            }
        }
    }
}
