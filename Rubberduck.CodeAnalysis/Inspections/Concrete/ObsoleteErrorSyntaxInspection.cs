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
    public sealed class ObsoleteErrorSyntaxInspection : ParseTreeInspectionBase
    {
        public ObsoleteErrorSyntaxInspection(RubberduckParserState state)
            : base(state)
        {
            Listener = new ObsoleteErrorSyntaxListener();
        }

        public override IInspectionListener Listener { get; }
        protected override string ResultDescription(QualifiedContext<ParserRuleContext> context)
        {
            return InspectionResults.ObsoleteErrorSyntaxInspection;
        }

        public class ObsoleteErrorSyntaxListener : InspectionListenerBase
        {
            public override void ExitErrorStmt(VBAParser.ErrorStmtContext context)
            {
                SaveContext(context);
            }
        }
    }
}
