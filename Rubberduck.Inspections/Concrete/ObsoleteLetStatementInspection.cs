using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class ObsoleteLetStatementInspection : ParseTreeInspectionBase
    {
        public ObsoleteLetStatementInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion)
        {
            Listener = new ObsoleteLetStatementListener();
        }

        public override CodeInspectionType InspectionType => CodeInspectionType.LanguageOpportunities;
        public override IInspectionListener Listener { get; }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return Listener.Contexts.Where(context => !IsIgnoringInspectionResultFor(context.ModuleName, context.Context.Start.Line))
                .Select(context => new QualifiedContextInspectionResult(this, InspectionsUI.ObsoleteLetStatementInspectionResultFormat, context));
        }

        public class ObsoleteLetStatementListener : VBAParserBaseListener, IInspectionListener
        {
            private readonly List<QualifiedContext<ParserRuleContext>> _contexts = new List<QualifiedContext<ParserRuleContext>>();
            public IReadOnlyList<QualifiedContext<ParserRuleContext>> Contexts => _contexts;

            public QualifiedModuleName CurrentModuleName { get; set; }

            public void ClearContexts()
            {
                _contexts.Clear();
            }

            public override void ExitLetStmt(VBAParser.LetStmtContext context)
            {
                if (context.LET() != null)
                {
                    _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
                }
            }
        }
    }
}
