using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Locates legacy 'Error' statements.
    /// </summary>
    /// <why>
    /// The legacy syntax is obsolete; prefer 'Err.Raise' instead.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Error 5 ' raises run-time error 5
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Err.Raise 5 ' raises run-time error 5
    /// End Sub
    /// ]]>
    /// </module>
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
