using Antlr4.Runtime.Tree;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete.ThunderCode
{
    /// <summary hidden="true">
    /// A ThunderCode inspection that locates instances of 'On Error GoTo -1' statements.
    /// </summary>
    /// <why>
    /// This inpection is flagging code we dubbed "ThunderCode", 
    /// code our friend Andrew Jackson would have written to confuse Rubberduck's parser and/or resolver. 
    /// 'On Error GoTo -1' is poorly documented and uselessly complicates error handling. Consider using 'On Error GoTo 0' instead.
    /// </why>
    internal sealed class OnErrorGoToMinusOneInspection : ParseTreeInspectionBase<VBAParser.OnErrorStmtContext>
    {
        public OnErrorGoToMinusOneInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
            ContextListener = new OnErrorGoToMinusOneListener();
        }

        protected override IInspectionListener<VBAParser.OnErrorStmtContext> ContextListener { get; }

        protected override string ResultDescription(QualifiedContext<VBAParser.OnErrorStmtContext> context)
        {
            return InspectionResults.OnErrorGoToMinusOneInspection.ThunderCodeFormat();
        }

        private class OnErrorGoToMinusOneListener : InspectionListenerBase<VBAParser.OnErrorStmtContext>
        {
            public override void EnterOnErrorStmt(VBAParser.OnErrorStmtContext context)
            {
                CheckContext(context, context.expression());
                base.EnterOnErrorStmt(context);
            }
            
            private void CheckContext(VBAParser.OnErrorStmtContext context, IParseTree expression)
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
