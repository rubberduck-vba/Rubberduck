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
    /// A ThunderCode inspection that locates instances of 'On Error GoTo -1' statements.
    /// </summary>
    /// <why>
    /// This inpection is flagging code we dubbed "ThunderCode", 
    /// code our friend Andrew Jackson would have written to confuse Rubberduck's parser and/or resolver. 
    /// 'On Error GoTo -1' is poorly documented and uselessly complicates error handling.
    /// </why>
    public class OnErrorGoToMinusOneInspection : ParseTreeInspectionBase
    {
        public OnErrorGoToMinusOneInspection(RubberduckParserState state) : base(state)
        {
            Listener = new OnErrorGoToMinusOneListener();
        }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return Listener.Contexts.Select(c => new QualifiedContextInspectionResult(
                this, InspectionResults.OnErrorGoToMinusOneInspection.ThunderCodeFormat(), c));
        }

        public override IInspectionListener Listener { get; }

        public class OnErrorGoToMinusOneListener : VBAParserBaseListener, IInspectionListener
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
            
            private void CheckContext(ParserRuleContext context, IParseTree expression)
            {
                var target = expression?.GetText().Trim() ?? string.Empty;
                if (target.StartsWith("-") && int.TryParse(target.Substring(1), out var result) && result == 1)
                {
                    _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
                }
            }
        }
    }
}
