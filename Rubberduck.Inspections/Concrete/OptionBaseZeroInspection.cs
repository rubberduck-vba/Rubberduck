using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class OptionBaseZeroInspection : InspectionBase, IParseTreeInspection
    {
        private IEnumerable<QualifiedContext> _parseTreeResults;

        public OptionBaseZeroInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Hint)
        {
        }

        public override string Meta => InspectionsUI.OptionBaseZeroInspectionMeta;
        public override string Description => InspectionsUI.OptionBaseZeroInspectionName;
        public override CodeInspectionType InspectionType => CodeInspectionType.MaintainabilityAndReadabilityIssues;

        public IEnumerable<QualifiedContext<VBAParser.OptionBaseStmtContext>> ParseTreeResults => _parseTreeResults.OfType<QualifiedContext<VBAParser.OptionBaseStmtContext>>();
        public void SetResults(IEnumerable<QualifiedContext> results) { _parseTreeResults = results; } 

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            if (ParseTreeResults == null)
            {
                return new InspectionResultBase[] { };
            }

            return ParseTreeResults.Where(context => !IsIgnoringInspectionResultFor(context.ModuleName.Component, context.Context.Start.Line))
                                   .Select(context => new OptionBaseZeroInspectionResult(this, new QualifiedContext<VBAParser.OptionBaseStmtContext>(context.ModuleName, context.Context)));
        }

        public class OptionBaseStatementListener : VBAParserBaseListener
        {
            private readonly IList<VBAParser.OptionBaseStmtContext> _contexts = new List<VBAParser.OptionBaseStmtContext>();
            public IEnumerable<VBAParser.OptionBaseStmtContext> Contexts => _contexts;

            public override void ExitOptionBaseStmt(VBAParser.OptionBaseStmtContext context)
            {
                if (context.numberLiteral()?.INTEGERLITERAL().Symbol.Text == "0")
                {
                    _contexts.Add(context);
                }
            }
        }
    }
}
