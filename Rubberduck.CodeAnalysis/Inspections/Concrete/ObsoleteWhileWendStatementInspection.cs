using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Flags 'While...Wend' loops as obsolete.
    /// </summary>
    /// <why>
    /// 'While...Wend' loops were made obsolete when 'Do While...Loop' statements were introduced.
    /// 'While...Wend' loops cannot be exited early without a GoTo jump; 'Do...Loop' statements can be conditionally exited with 'Exit Do'.
    /// </why>
    /// <example hasresult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     While True
    ///         ' ...
    ///     Wend
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasresult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Do While True
    ///         ' ...
    ///     Loop
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class ObsoleteWhileWendStatementInspection : ParseTreeInspectionBase<VBAParser.WhileWendStmtContext>
    {
        public ObsoleteWhileWendStatementInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
            ContextListener = new ObsoleteWhileWendStatementListener();
        }

        protected override IInspectionListener<VBAParser.WhileWendStmtContext> ContextListener { get; }

        protected override string ResultDescription(QualifiedContext<VBAParser.WhileWendStmtContext> context)
        {
            return InspectionResults.ObsoleteWhileWendStatementInspection;
        }

        private class ObsoleteWhileWendStatementListener : InspectionListenerBase<VBAParser.WhileWendStmtContext>
        {
            public override void ExitWhileWendStmt(VBAParser.WhileWendStmtContext context)
            {
                SaveContext(context);
            }
        }
    }
}