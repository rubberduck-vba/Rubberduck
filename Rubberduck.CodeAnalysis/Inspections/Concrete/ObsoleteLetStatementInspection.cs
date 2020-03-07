using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Locates explicit 'Let' assignments.
    /// </summary>
    /// <why>
    /// The legacy syntax is obsolete/redundant; prefer implicit Let-coercion instead.
    /// </why>
    /// <example hasResult="true">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim foo As Long
    ///     Let foo = 42 ' explicit Let is redundant
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResult="false">
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
