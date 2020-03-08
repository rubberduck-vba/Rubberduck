using Antlr4.Runtime.Misc;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Locates 'For' loops where the 'Step' token is specified with the default increment value (1).
    /// </summary>
    /// <why>
    /// Out of convention or preference, explicit 'Step 1' specifiers could be considered redundant; 
    /// this inspection can ensure the consistency of the convention.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim i As Long
    ///     For i = 1 To 100 Step 1 ' 1 being the implicit default, 'Step 1' could be considered redundant.
    ///         ' ...
    ///     Next
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim i As Long
    ///     For i = 1 To 100 ' implicit: 'Step 1'
    ///         ' ...
    ///     Next
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class StepOneIsRedundantInspection : ParseTreeInspectionBase<VBAParser.StepStmtContext>
    {
        public StepOneIsRedundantInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
            ContextListener = new StepOneIsRedundantListener();
        }

        protected override IInspectionListener<VBAParser.StepStmtContext> ContextListener { get; }

        protected override string ResultDescription(QualifiedContext<VBAParser.StepStmtContext> context)
        {
            return InspectionResults.StepOneIsRedundantInspection;
        }

        private class StepOneIsRedundantListener : InspectionListenerBase<VBAParser.StepStmtContext>
        {
            public override void EnterForNextStmt([NotNull] VBAParser.ForNextStmtContext context)
            {
                var stepStatement = context.stepStmt();

                if (stepStatement == null)
                {
                    return;
                }

                var stepText = stepStatement.expression().GetText();

                if (stepText == "1")
                {
                    SaveContext(stepStatement);
                }
            }
        }
    }
}
