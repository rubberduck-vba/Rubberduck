using Rubberduck.Inspections.Abstract;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Inspections.Extensions;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;
using Rubberduck.VBEditor;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Flags 'While...Wend' loops as obsolete.
    /// </summary>
    /// <why>
    /// 'While...Wend' loops were made obsolete when 'Do While...Loop' statements were introduced.
    /// 'While...Wend' loops cannot be exited early without a GoTo jump; 'Do...Loop' statements can be conditionally exited with 'Exit Do'.
    /// </why>
    /// <example hasresult="true">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     While True
    ///         ' ...
    ///     Wend
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasresult="false">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Do While True
    ///         ' ...
    ///     Loop
    /// End Sub
    /// ]]>
    /// </example>
    public sealed class ObsoleteWhileWendStatementInspection : ParseTreeInspectionBase
    {
        public ObsoleteWhileWendStatementInspection(RubberduckParserState state)
            : base(state)
        {
            Listener = new ObsoleteWhileWendStatementListener();
        }

        public override IInspectionListener Listener { get; }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return Listener.Contexts
                .Select(context => new QualifiedContextInspectionResult(this, InspectionResults.ObsoleteWhileWendStatementInspection, context));
        }

        public class ObsoleteWhileWendStatementListener : VBAParserBaseListener, IInspectionListener
        {
            private readonly List<QualifiedContext<ParserRuleContext>> _contexts =
                new List<QualifiedContext<ParserRuleContext>>();

            public IReadOnlyList<QualifiedContext<ParserRuleContext>> Contexts => _contexts;

            public QualifiedModuleName CurrentModuleName { get; set; }

            public void ClearContexts()
            {
                _contexts.Clear();
            }

            public override void ExitWhileWendStmt(VBAParser.WhileWendStmtContext context)
            {
                _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
            }
        }
    }
}