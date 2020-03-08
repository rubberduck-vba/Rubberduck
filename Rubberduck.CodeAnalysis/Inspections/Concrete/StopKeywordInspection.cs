using Antlr4.Runtime.Misc;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Locates 'Stop' instructions in user code.
    /// </summary>
    /// <why>
    /// While a great debugging tool, 'Stop' instructions should not be reachable in production code; this inspection makes it easy to locate them all.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     ' ...
    ///     Stop ' halts execution on-the-spot, bringing up the VBE; not very user-friendly!
    ///     '....
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     ' ...
    ///     'Stop ' the commented-out statement isn't executable. Could also be simply removed.
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class StopKeywordInspection : ParseTreeInspectionBase<VBAParser.StopStmtContext>
    {
        public StopKeywordInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
            ContextListener = new StopKeywordListener();
        }

        protected override IInspectionListener<VBAParser.StopStmtContext> ContextListener { get; }

        protected override string ResultDescription(QualifiedContext<VBAParser.StopStmtContext> context)
        {
            return InspectionResults.StopKeywordInspection;
        }

        private class StopKeywordListener : InspectionListenerBase<VBAParser.StopStmtContext>
        {
            public override void ExitStopStmt([NotNull] VBAParser.StopStmtContext context)
            {
                SaveContext(context);
            }
        }
    }
}
