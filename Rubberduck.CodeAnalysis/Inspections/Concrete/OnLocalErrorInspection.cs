using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;
using Antlr4.Runtime.Misc;
using Rubberduck.CodeAnalysis.Inspections;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Flags obsolete 'On Local Error' statements.
    /// </summary>
    /// <why>
    /// All errors are "local" - the keyword is redundant/confusing and should be removed.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     On Local Error GoTo ErrHandler
    ///     ' ...
    ///     Exit Sub
    /// ErrHandler:
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     On Error GoTo ErrHandler
    ///     ' ...
    ///     Exit Sub
    /// ErrHandler:
    ///     ' ...
    /// End Sub
    /// ]]>
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
