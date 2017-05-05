using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
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
    public sealed class MultipleDeclarationsInspection : InspectionBase, IParseTreeInspection
    {
        public MultipleDeclarationsInspection(RubberduckParserState state)
            : base(state) { }

        public override CodeInspectionType InspectionType => CodeInspectionType.MaintainabilityAndReadabilityIssues;

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            return Listener.Contexts
                .Where(result => !IsIgnoringInspectionResultFor(result.ModuleName, result.Context.Start.Line))
                .Select(context => new QualifiedContextInspectionResult(this,
                                                        InspectionsUI.MultipleDeclarationsInspectionResultFormat,
                                                        State,
                                                        context));
        }

        public IInspectionListener Listener { get; } = new ParameterListListener();

        public class ParameterListListener : VBAParserBaseListener, IInspectionListener
        {
            private readonly List<QualifiedContext<ParserRuleContext>> _contexts = new List<QualifiedContext<ParserRuleContext>>();
            public IReadOnlyList<QualifiedContext<ParserRuleContext>> Contexts => _contexts;

            public QualifiedModuleName CurrentModuleName { get; set; }

            public void ClearContexts()
            {
                _contexts.Clear();
            }

            public override void ExitVariableListStmt([NotNull] VBAParser.VariableListStmtContext context)
            {
                if (context.variableSubStmt().Count > 1)
                {
                    _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
                }
            }

            public override void ExitConstStmt([NotNull] VBAParser.ConstStmtContext context)
            {
                if (context.constSubStmt().Count > 1)
                {
                    _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
                }
            }
        }
    }
}
