using Antlr4.Runtime.Misc;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Flags obsolete 'On Local Error' statements.
    /// </summary>
    /// <why>
    /// All errors are "local" - the keyword is redundant/confusing and should be removed.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     On Local Error GoTo ErrHandler
    ///     ' ...
    ///     Exit Sub
    /// ErrHandler:
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     On Error GoTo ErrHandler
    ///     ' ...
    ///     Exit Sub
    /// ErrHandler:
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class OnLocalErrorInspection : ParseTreeInspectionBase<VBAParser.OnErrorStmtContext>
    {
        public OnLocalErrorInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
            ContextListener = new OnLocalErrorListener();
        }

        protected override IInspectionListener<VBAParser.OnErrorStmtContext> ContextListener { get; }

        protected override string ResultDescription(QualifiedContext<VBAParser.OnErrorStmtContext> context)
        {
            return InspectionResults.OnLocalErrorInspection;
        }

        private class OnLocalErrorListener : InspectionListenerBase<VBAParser.OnErrorStmtContext>
        {
            public override void ExitOnErrorStmt([NotNull] VBAParser.OnErrorStmtContext context)
            {
                if (context.ON_LOCAL_ERROR() != null)
                {
                   SaveContext(context);
                }
            }
        }
    }
}
