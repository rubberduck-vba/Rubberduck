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
    public sealed class ModuleScopeDimKeywordInspection : InspectionBase, IParseTreeInspection
    {
        public ModuleScopeDimKeywordInspection(RubberduckParserState state) 
            : base(state, CodeInspectionSeverity.Suggestion) { }

        public override CodeInspectionType InspectionType => CodeInspectionType.MaintainabilityAndReadabilityIssues;

        public IInspectionListener Listener { get; } = new ModuleScopedDimListener();

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            return Listener.Contexts
                .Where(result => !IsIgnoringInspectionResultFor(result.ModuleName, result.Context.Start.Line))
                .SelectMany(result => result.Context.FindChildren<VBAParser.VariableSubStmtContext>()
                        .Select(r => new QualifiedContext<ParserRuleContext>(result.ModuleName, r)))
                .Select(result => new QualifiedContextInspectionResult(this,
                                                       string.Format(InspectionsUI.ModuleScopeDimKeywordInspectionResultFormat, ((VBAParser.VariableSubStmtContext)result.Context).identifier().GetText()),
                                                       State,
                                                       result));
        }

        public class ModuleScopedDimListener : VBAParserBaseListener, IInspectionListener
        {
            private readonly List<QualifiedContext<ParserRuleContext>> _contexts = new List<QualifiedContext<ParserRuleContext>>();
            public IReadOnlyList<QualifiedContext<ParserRuleContext>> Contexts => _contexts;

            public QualifiedModuleName CurrentModuleName { get; set; }

            public void ClearContexts()
            {
                _contexts.Clear();
            }

            public override void ExitVariableStmt([NotNull] VBAParser.VariableStmtContext context)
            {
                if (context.DIM() != null && context.Parent is VBAParser.ModuleDeclarationsElementContext)
                {
                    _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
                }
            }
        }
    }
}