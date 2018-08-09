using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Inspections.Concrete
{
    public sealed class ObsoleteCallingConventionInspection : ParseTreeInspectionBase
    {
        public ObsoleteCallingConventionInspection(RubberduckParserState state)
            : base(state)
        {
            Listener = new ObsoleteCallingConventionListener();
        }

        public override IInspectionListener Listener { get; }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return Listener.Contexts
                .Where(context => ((VBAParser.DeclareStmtContext) context.Context).CDECL() != null &&
                    !IsIgnoringInspectionResultFor(context.ModuleName, context.Context.Start.Line))
                .Select(context => new QualifiedContextInspectionResult(this,
                    string.Format(InspectionResults.ObsoleteCallingConventionInspection,
                        ((VBAParser.DeclareStmtContext) context.Context).identifier().GetText()), context));
        }

        public class ObsoleteCallingConventionListener : VBAParserBaseListener, IInspectionListener
        {
            private readonly List<QualifiedContext<ParserRuleContext>> _contexts = new List<QualifiedContext<ParserRuleContext>>();
            public IReadOnlyList<QualifiedContext<ParserRuleContext>> Contexts => _contexts;

            public QualifiedModuleName CurrentModuleName { get; set; }

            public void ClearContexts()
            {
                _contexts.Clear();
            }

            public override void ExitDeclareStmt(VBAParser.DeclareStmtContext context)
            {
                _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
                base.ExitDeclareStmt(context);
            }
        }
    }
}
