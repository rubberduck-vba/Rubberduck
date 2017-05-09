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
    public sealed class ObsoleteCallStatementInspection : InspectionBase, IParseTreeInspection
    {
        public ObsoleteCallStatementInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion) { }

        public override CodeInspectionType InspectionType => CodeInspectionType.LanguageOpportunities;

        public IInspectionListener Listener { get; } =
            new ObsoleteCallStatementListener();

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            var results = new List<IInspectionResult>();

            foreach (var context in Listener.Contexts.Where(context => !IsIgnoringInspectionResultFor(context.ModuleName, context.Context.Start.Line)))
            {
                var module = context.ModuleName.Component.CodeModule;
                var lines = module.GetLines(context.Context.Start.Line,
                    context.Context.Stop.Line - context.Context.Start.Line + 1);

                var stringStrippedLines = string.Join(string.Empty, lines).StripStringLiterals();

                int commentIndex;
                if (stringStrippedLines.HasComment(out commentIndex))
                {
                    stringStrippedLines = stringStrippedLines.Remove(commentIndex);
                }

                if (!stringStrippedLines.Contains(":"))
                {
                    results.Add(new QualifiedContextInspectionResult(this,
                                                     InspectionsUI.ObsoleteCallStatementInspectionResultFormat,
                                                     State,
                                                     context));
                }
            }

            return results;
        }

        public class ObsoleteCallStatementListener : VBAParserBaseListener, IInspectionListener
        {
            private readonly List<QualifiedContext<ParserRuleContext>> _contexts = new List<QualifiedContext<ParserRuleContext>>();
            public IReadOnlyList<QualifiedContext<ParserRuleContext>> Contexts => _contexts;

            public QualifiedModuleName CurrentModuleName { get; set; }

            public void ClearContexts()
            {
                _contexts.Clear();
            }

            public override void ExitCallStmt(VBAParser.CallStmtContext context)
            {
                if (context.CALL() != null)
                {
                    _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
                }
            }
        }
    }
}
