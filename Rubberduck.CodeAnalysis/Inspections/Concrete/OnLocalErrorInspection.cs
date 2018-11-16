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
using Antlr4.Runtime.Misc;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class OnLocalErrorInspection : ParseTreeInspectionBase
    {
        public OnLocalErrorInspection(RubberduckParserState state)
            : base(state) { }

        public override IInspectionListener Listener { get; } =
            new OnLocalErrorListener();

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return Listener.Contexts
                .Where(result => !IsIgnoringInspectionResultFor(result.ModuleName, result.Context.Start.Line))
                .Select(result => new QualifiedContextInspectionResult(this,
                                                       InspectionResults.OnLocalErrorInspection,
                                                       result));
        }

        public class OnLocalErrorListener : VBAParserBaseListener, IInspectionListener
        {
            private readonly List<QualifiedContext<ParserRuleContext>> _contexts = new List<QualifiedContext<ParserRuleContext>>();
            public IReadOnlyList<QualifiedContext<ParserRuleContext>> Contexts => _contexts;
            
            public QualifiedModuleName CurrentModuleName { get; set; }

            public void ClearContexts()
            {
                _contexts.Clear();
            }

            public override void ExitOnErrorStmt([NotNull] VBAParser.OnErrorStmtContext context)
            {
                if (context.ON_LOCAL_ERROR() != null)
                {
                    _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
                }
            }
        }
    }
}
