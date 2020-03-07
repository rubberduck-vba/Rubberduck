using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Locates explicit 'Let' assignments.
    /// </summary>
    /// <why>
    /// The legacy syntax is obsolete/redundant; prefer implicit Let-coercion instead.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim foo As Long
    ///     Let foo = 42 ' explicit Let is redundant
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim foo As Long
    ///     foo = 42 ' [Let] is implicit
    /// End Sub
    /// ]]>
    /// </example>
    internal sealed class ObsoleteLetStatementInspection : ParseTreeInspectionBase<VBAParser.LetStmtContext>
    {
        public ObsoleteLetStatementInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
            ContextListener = new ObsoleteLetStatementListener();
        }
        
        protected override IInspectionListener<VBAParser.LetStmtContext> ContextListener { get; }

        protected override string ResultDescription(QualifiedContext<VBAParser.LetStmtContext> context)
        {
            return InspectionResults.ObsoleteLetStatementInspection;
        }

        private class ObsoleteLetStatementListener : InspectionListenerBase<VBAParser.LetStmtContext>
        {
            public override void ExitLetStmt(VBAParser.LetStmtContext context)
            {
                if (context.LET() != null)
                {
                   SaveContext(context);
                }
            }
        }
    }
}
