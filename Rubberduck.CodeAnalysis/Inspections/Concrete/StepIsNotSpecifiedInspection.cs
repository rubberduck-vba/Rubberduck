using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;
using Antlr4.Runtime.Misc;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Locates 'For' loops where the 'Step' token is omitted.
    /// </summary>
    /// <why>
    /// Out of convention or preference, explicit 'Step' specifiers could be considered mandatory; 
    /// this inspection can ensure the consistency of the convention.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim i As Long
    ///     For i = 1 To 100 ' Step is implicitly 1
    ///         ' ...
    ///     Next
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim i As Long
    ///     For i = 1 To 100 Step 1 ' explicit 'Step 1' could also be considered redundant.
    ///         ' ...
    ///     Next
    /// End Sub
    /// ]]>
    /// </example>
    public sealed class StepIsNotSpecifiedInspection : ParseTreeInspectionBase
    {
        public StepIsNotSpecifiedInspection(RubberduckParserState state) 
            : base(state) { }

        protected override string ResultDescription(QualifiedContext<ParserRuleContext> context)
        {
            return InspectionResults.StepIsNotSpecifiedInspection;
        }

        public override IInspectionListener Listener { get; } =
            new StepIsNotSpecifiedListener();
    }

    public class StepIsNotSpecifiedListener : InspectionListenerBase
    {
        public override void EnterForNextStmt([NotNull] VBAParser.ForNextStmtContext context)
        {
            var stepStatement = context.stepStmt();

            if (stepStatement == null)
            {
                SaveContext(context);
            }
        }
    }
}
