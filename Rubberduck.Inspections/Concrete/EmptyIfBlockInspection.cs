using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    internal class EmptyIfBlockInspection : EmptyBlockInspectionBase<EmptyIfBlockInspection>
    {
        public EmptyIfBlockInspection(RubberduckParserState state) 
            : base(state, InspectionsUI.EmptyIfBlockInspectionResultFormat) { }

        public override IInspectionListener Listener { get; } =
            new EmptyIfBlockListener();

        public class EmptyIfBlockListener : EmptyBlockListenerBase
        {
            public override void EnterIfStmt([NotNull] VBAParser.IfStmtContext context)
            {
                InspectBlockForExecutableStatements(context.block(), context);
            }

            public override void EnterElseIfBlock([NotNull] VBAParser.ElseIfBlockContext context)
            {
                InspectBlockForExecutableStatements(context.block(), context);
            }

            public override void EnterSingleLineIfStmt([NotNull] VBAParser.SingleLineIfStmtContext context)
            {
                if (context.ifWithEmptyThen() != null)
                {
                    AddResult(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context.ifWithEmptyThen()));
                }
            }
        }
    }
}
