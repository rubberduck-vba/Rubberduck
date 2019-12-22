using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Inspections.Concrete.ThunderCode
{
    /// <summary hidden="true">
    /// A ThunderCode inspection that locates negative line numbers.
    /// </summary>
    /// <why>
    /// This inpection is flagging code we dubbed "ThunderCode", 
    /// code our friend Andrew Jackson would have written to confuse Rubberduck's parser and/or resolver. 
    /// The VBE does allow rather strange and unbelievable things to happen.
    /// </why>
    public class NegativeLineNumberInspection : ParseTreeInspectionBase
    {
        public NegativeLineNumberInspection(RubberduckParserState state) : base(state)
        {
            Listener = new NegativeLineNumberKeywordsListener();
        }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return Listener.Contexts.Select(c => new QualifiedContextInspectionResult(
                this, InspectionResults.NegativeLineNumberInspection.ThunderCodeFormat(), c));
        }

        public override IInspectionListener Listener { get; }

        public class NegativeLineNumberKeywordsListener : VBAParserBaseListener, IInspectionListener
        {
            private readonly List<QualifiedContext<ParserRuleContext>> _contexts = new List<QualifiedContext<ParserRuleContext>>();

            public IReadOnlyList<QualifiedContext<ParserRuleContext>> Contexts => _contexts;

            public void ClearContexts() => _contexts.Clear();

            public QualifiedModuleName CurrentModuleName { get; set; }

            public override void EnterOnErrorStmt(VBAParser.OnErrorStmtContext context)
            {
                CheckContext(context, context.expression());
                base.EnterOnErrorStmt(context);
            }

            public override void EnterGoToStmt(VBAParser.GoToStmtContext context)
            {
                CheckContext(context, context.expression());
                base.EnterGoToStmt(context);
            }

            public override void EnterLineNumberLabel(VBAParser.LineNumberLabelContext context)
            {
                CheckContext(context, context);
                base.EnterLineNumberLabel(context);
            }

            private void CheckContext(ParserRuleContext context, IParseTree expression)
            {
                var target = expression?.GetText().Trim() ?? string.Empty;
                if (target.StartsWith("-") && int.TryParse(target.Substring(1), out _))
                {
                    _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
                }
            }
        }
    }
}
