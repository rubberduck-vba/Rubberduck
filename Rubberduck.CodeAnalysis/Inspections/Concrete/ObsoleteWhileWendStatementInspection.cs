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
            return Listener.Contexts.Where(context =>
                    !context.IsIgnoringInspectionResultFor(State.DeclarationFinder, AnnotationName))
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