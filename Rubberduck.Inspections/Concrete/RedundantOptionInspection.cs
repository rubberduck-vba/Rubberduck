using System;
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
    public sealed class RedundantOptionInspection : ParseTreeInspectionBase
    {
        public RedundantOptionInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Hint)
        {
            Listener = new RedundantModuleOptionListener();
        }

        public override string Meta => InspectionsUI.RedundantOptionInspectionMeta;
        public override string Description => InspectionsUI.RedundantOptionInspectionName;
        public override CodeInspectionType InspectionType => CodeInspectionType.LanguageOpportunities;

        public override IInspectionListener Listener { get; }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return Listener.Contexts.Where(context => !IsIgnoringInspectionResultFor(context.ModuleName, context.Context.Start.Line))
                                   .Select(context => new QualifiedContextInspectionResult(this,
                                                                           string.Format(InspectionsUI.RedundantOptionInspectionResultFormat, context.Context.GetText()),
                                                                           context));
        }

        public class RedundantModuleOptionListener : VBAParserBaseListener, IInspectionListener
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
                if (context.numberLiteral()?.INTEGERLITERAL().Symbol.Text == "0")
                {
                    _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
                }
            }

            public override void ExitOptionCompareStmt(VBAParser.OptionCompareStmtContext context)
            {
                // BINARY is the default, and DATABASE is specified by default + only valid in Access.
                if (context.TEXT() == null && context.DATABASE() == null)
                {
                    _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
                }
            }
        }
    }
}
