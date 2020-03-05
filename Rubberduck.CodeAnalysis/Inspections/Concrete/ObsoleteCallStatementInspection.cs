using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Locates explicit 'Call' statements.
    /// </summary>
    /// <why>
    /// The 'Call' keyword is obsolete and redundant, since call statements are legal and generally more consistent without it.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Public Sub Test()
    ///     Call DoSomething(42)
    /// End Sub
    ///
    /// Private Sub DoSomething(ByVal foo As Long)
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Public Sub Test()
    ///     DoSomething 42
    /// End Sub
    ///
    /// Private Sub DoSomething(ByVal foo As Long)
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </example>
    public sealed class ObsoleteCallStatementInspection : ParseTreeInspectionBase<VBAParser.CallStmtContext>
    {
        public ObsoleteCallStatementInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
            ContextListener = new ObsoleteCallStatementListener();
        }

        protected override IInspectionListener<VBAParser.CallStmtContext> ContextListener { get; }

        protected override string ResultDescription(QualifiedContext<VBAParser.CallStmtContext> context)
        {
            return InspectionResults.ObsoleteCallStatementInspection;
        }

        protected override bool IsResultContext(QualifiedContext<VBAParser.CallStmtContext> context, DeclarationFinder finder)
        {
            if (!context.Context.TryGetFollowingContext(out VBAParser.IndividualNonEOFEndOfStatementContext followingEndOfStatement)
                || followingEndOfStatement.COLON() == null)
            {
                return true;
            }

            if (!context.Context.TryGetPrecedingContext(out VBAParser.IndividualNonEOFEndOfStatementContext precedingEndOfStatement)
                || precedingEndOfStatement.endOfLine() == null)
            {
                return true;
            }

            return false;
        }

        public class ObsoleteCallStatementListener : InspectionListenerBase<VBAParser.CallStmtContext>
        {
            public override void ExitCallStmt(VBAParser.CallStmtContext context)
            {
                if (context.CALL() != null)
                {
                    SaveContext(context);
                }
            }
        }
    }
}
