using System;
using System.Text;
using Antlr4.Runtime;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete
{
    /// <summary>
    /// Splits multiple declarations into separate statements.
    /// </summary>
    /// <inspections>
    /// <inspection name="MultipleDeclarationsInspection" />
    /// </inspections>
    /// <canfix multiple="true" procedure="true" module="true" project="true" all="true" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     Dim value As Long, something As String
    ///     '...
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     Dim value As Long
    ///     Dim something As String
    ///     '...
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    internal sealed class SplitMultipleDeclarationsQuickFix : QuickFixBase
    {
        public SplitMultipleDeclarationsQuickFix()
            : base(typeof(MultipleDeclarationsInspection))
        {}

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var context = result.Context is VBAParser.ConstStmtContext
                ? result.Context
                : (ParserRuleContext)result.Context.Parent;

            string declarationsText;
            switch (context)
            {
                case VBAParser.ConstStmtContext consts:
                    declarationsText = GetDeclarationsText(consts);
                    break;
                case VBAParser.VariableStmtContext variables:
                    declarationsText = GetDeclarationsText(variables);
                    break;
                default:
                    throw new NotSupportedException();
            }

            var rewriter = rewriteSession.CheckOutModuleRewriter(result.QualifiedSelection.QualifiedName);
            rewriter.Replace(context, declarationsText);
        }

        private string GetDeclarationsText(VBAParser.ConstStmtContext consts)
        {
            var keyword = string.Empty;
            if (consts.visibility() != null)
            {
                keyword += consts.visibility().GetText() + ' ';
            }

            keyword += consts.CONST().GetText() + ' ';

            var newContent = new StringBuilder();
            foreach (var constant in consts.constSubStmt())
            {
                newContent.AppendLine(keyword + constant.GetText());
            }

            return newContent.ToString();
        }

        private string GetDeclarationsText(VBAParser.VariableStmtContext variables)
        {
            var keyword = string.Empty;
            if (variables.DIM() != null)
            {
                keyword += Tokens.Dim + ' ';
            }
            else if (variables.visibility() != null)
            {
                keyword += variables.visibility().GetText() + ' ';
            }
            else if (variables.STATIC() != null)
            {
                keyword += variables.STATIC().GetText() + ' ';
            }

            var newContent = new StringBuilder();
            foreach (var variable in variables.variableListStmt().variableSubStmt())
            {
                newContent.AppendLine(keyword + variable.GetText());
            }

            return newContent.ToString();
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.SplitMultipleDeclarationsQuickFix;

        public override bool CanFixMultiple => true;
        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
        public override bool CanFixAll => true;
    }
}