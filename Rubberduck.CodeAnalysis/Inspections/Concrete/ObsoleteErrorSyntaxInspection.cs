using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.Inspections.Inspections.Extensions;

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

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return Listener.Contexts
                .Select(context => new QualifiedContextInspectionResult(this, InspectionResults.ObsoleteErrorSyntaxInspection, context));
        }

        public class ObsoleteErrorSyntaxListener : VBAParserBaseListener, IInspectionListener
        {
            private readonly List<QualifiedContext<ParserRuleContext>> _contexts = new List<QualifiedContext<ParserRuleContext>>();
            public IReadOnlyList<QualifiedContext<ParserRuleContext>> Contexts => _contexts;

            public QualifiedModuleName CurrentModuleName { get; set; }

            public void ClearContexts()
            {
                _contexts.Clear();
            }

            public override void ExitErrorStmt(VBAParser.ErrorStmtContext context)
            {
                _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
            }
        }
    }
}
