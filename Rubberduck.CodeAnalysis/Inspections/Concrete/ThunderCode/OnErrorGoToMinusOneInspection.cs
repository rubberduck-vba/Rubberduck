using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;

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
        public OnErrorGoToMinusOneInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
            Listener = new OnErrorGoToMinusOneListener();
        }

        protected override string ResultDescription(QualifiedContext<ParserRuleContext> context)
        {
            return InspectionResults.OnErrorGoToMinusOneInspection.ThunderCodeFormat();
        }

        public override IInspectionListener Listener { get; }

        public class OnErrorGoToMinusOneListener : InspectionListenerBase
        {
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
                   SaveContext(context);
                }
            }
        }
    }
}
