using Antlr4.Runtime.Misc;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Common;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;
using Antlr4.Runtime;
using Rubberduck.Resources.Experimentals;
using Rubberduck.Parsing;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Identifies empty 'For...Next' blocks that can be safely removed.
    /// </summary>
    /// <why>
    /// Dead code should be removed. A loop without a body is usually redundant.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal foo As Long)
    ///     Dim i As Long
    ///     For i = 0 To foo
    ///         ' no executable statement...
    ///     Next
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal foo As Long)
    ///     Dim i As Long
    ///     For i = 0 To foo
    ///         Debug.Print i
    ///     Next
    /// End Sub
    /// ]]>
    /// </example>
    [Experimental(nameof(ExperimentalNames.EmptyBlockInspections))]
    internal class EmptyForLoopBlockInspection : ParseTreeInspectionBase
    {
        public EmptyForLoopBlockInspection(RubberduckParserState state)
            : base(state) { }

        protected override string ResultDescription(QualifiedContext<ParserRuleContext> context)
        {
            return InspectionResults.EmptyForLoopBlockInspection;
        }

        public override IInspectionListener Listener { get; } =
            new EmptyForLoopBlockListener();

        public class EmptyForLoopBlockListener : EmptyBlockInspectionListenerBase
        {
            public override void EnterForNextStmt([NotNull] VBAParser.ForNextStmtContext context)
            {
                InspectBlockForExecutableStatements(context.unterminatedBlock(), context);
            }
        }
    }
}
