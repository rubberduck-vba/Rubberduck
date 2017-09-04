using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;
using Antlr4.Runtime.Misc;

namespace Rubberduck.Inspections.Concrete
{
    internal class EmptyElseBlockInspection : EmptyBlockInspectionBase<EmptyElseBlockInspection>
    {
        public EmptyElseBlockInspection(RubberduckParserState state) 
            : base(state, InspectionsUI.EmptyElseBlockInspectionResultFormat) { }

        public override IInspectionListener Listener { get; } = new EmptyElseBlockListener();
        
        public class EmptyElseBlockListener : EmptyBlockListenerBase
        {
            public override void EnterElseBlock([NotNull] VBAParser.ElseBlockContext context)
            {
                InspectBlockForExecutableStatements(context.block(), context);
            }
        }
    }
}