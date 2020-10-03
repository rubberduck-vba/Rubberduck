using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete
{
    /// <summary>
    /// Removes an empty Else block.
    /// </summary>
    /// <inspections>
    /// <inspection name="EmptyElseBlockInspection" />
    /// </inspections>
    /// <canfix multiple="false" procedure="false" module="false" project="false" all="false" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     If Application.Calculation = xlCalculationAutomatic Then
    ///         Application.Calculation = xlCalculationManual
    ///     Else
    ///     End If
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     If Application.Calculation = xlCalculationAutomatic Then
    ///         Application.Calculation = xlCalculationManual
    ///     End If
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    internal sealed class RemoveEmptyElseBlockQuickFix : QuickFixBase
    {
        public RemoveEmptyElseBlockQuickFix()
            : base(typeof(EmptyElseBlockInspection))
        {}

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(result.QualifiedSelection.QualifiedName);

            UpdateContext((VBAParser.ElseBlockContext)result.Context, rewriter);
        }

        private void UpdateContext(VBAParser.ElseBlockContext context, IModuleRewriter rewriter)
        {
            var elseBlock = context.block();

            if (elseBlock.ChildCount == 0 )
            {
                rewriter.Remove(context);
            }
        }
        
        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.RemoveEmptyElseBlockQuickFix;

        public override bool CanFixMultiple => false;
        public override bool CanFixInProcedure => false;
        public override bool CanFixInModule => false;
        public override bool CanFixInProject => false;
        public override bool CanFixAll => false;
    }
}
