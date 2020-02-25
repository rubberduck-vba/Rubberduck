using Antlr4.Runtime;
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
    public sealed class ObsoleteLetStatementInspection : ParseTreeInspectionBase
    {
        public ObsoleteLetStatementInspection(RubberduckParserState state)
            : base(state)
        {
            Listener = new ObsoleteLetStatementListener();
        }
        
        public override IInspectionListener Listener { get; }

        protected override string ResultDescription(QualifiedContext<ParserRuleContext> context)
        {
            return InspectionResults.ObsoleteLetStatementInspection;
        }

        public class ObsoleteLetStatementListener : InspectionListenerBase
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
