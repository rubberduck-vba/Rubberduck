using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Extensions;
using Rubberduck.VBEditor.ComManagement;

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
    public sealed class ObsoleteCallStatementInspection : ParseTreeInspectionBase
    {
        public ObsoleteCallStatementInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
            Listener = new ObsoleteCallStatementListener();
        }

        public override IInspectionListener Listener { get; }
        protected override string ResultDescription(QualifiedContext<ParserRuleContext> context)
        {
            return InspectionResults.ObsoleteCallStatementInspection;
        }

        protected override bool IsResultContext(QualifiedContext<ParserRuleContext> context)
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

        public class ObsoleteCallStatementListener : InspectionListenerBase
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
