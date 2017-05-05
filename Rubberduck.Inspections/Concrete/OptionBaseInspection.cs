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
    public sealed class OptionBaseInspection : InspectionBase, IParseTreeInspection
    {
        public OptionBaseInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Hint) { }

        public override CodeInspectionType InspectionType => CodeInspectionType.MaintainabilityAndReadabilityIssues;

        public IInspectionListener Listener { get; } =
            new OptionBaseStatementListener();

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            return Listener.Contexts.Where(context => !IsIgnoringInspectionResultFor(context.ModuleName, context.Context.Start.Line))
                .Select(context => new QualifiedContextInspectionResult(this,
                                                        string.Format(InspectionsUI.OptionBaseInspectionResultFormat, context.ModuleName.ComponentName),
                                                        State,
                                                        context));
        }

        public class OptionBaseStatementListener : VBAParserBaseListener, IInspectionListener
        {
            private readonly List<QualifiedContext<ParserRuleContext>> _contexts = new List<QualifiedContext<ParserRuleContext>>();
            public IReadOnlyList<QualifiedContext<ParserRuleContext>> Contexts => _contexts;

            public QualifiedModuleName CurrentModuleName { get; set; }

            public void ClearContexts()
            {
                _contexts.Clear();
            }

            public override void ExitOptionBaseStmt(VBAParser.OptionBaseStmtContext context)
            {
                if (context.numberLiteral()?.INTEGERLITERAL().Symbol.Text == "1")
                {
                    _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
                }
            }
        }
    }
}
