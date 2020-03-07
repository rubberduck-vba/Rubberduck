using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Locates legacy 'Error' statements.
    /// </summary>
    /// <why>
    /// The legacy syntax is obsolete; prefer 'Err.Raise' instead.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Error 5 ' raises run-time error 5
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Err.Raise 5 ' raises run-time error 5
    /// End Sub
    /// ]]>
    /// </example>
    internal sealed class ObsoleteErrorSyntaxInspection : ParseTreeInspectionBase<VBAParser.ErrorStmtContext>
    {
        public ObsoleteErrorSyntaxInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
            ContextListener = new ObsoleteErrorSyntaxListener();
        }

        protected override IInspectionListener<VBAParser.ErrorStmtContext> ContextListener { get; }

        protected override string ResultDescription(QualifiedContext<VBAParser.ErrorStmtContext> context)
        {
            return InspectionResults.ObsoleteErrorSyntaxInspection;
        }

        private class ObsoleteErrorSyntaxListener : InspectionListenerBase<VBAParser.ErrorStmtContext>
        {
            public override void ExitErrorStmt(VBAParser.ErrorStmtContext context)
            {
                SaveContext(context);
            }
        }
    }
}
