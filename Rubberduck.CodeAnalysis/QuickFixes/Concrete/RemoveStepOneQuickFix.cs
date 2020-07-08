using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete
{
    /// <summary>
    /// Removes 'Step 1' specifier from 'For...Next' loop statement, 1 being the implicit default 'Step' increment.
    /// </summary>
    /// <inspections>
    /// <inspection name="StepOneIsRedundantInspection" />
    /// </inspections>
    /// <canfix multiple="true" procedure="true" module="true" project="true" all="true" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     Dim i As Long
    ///     For i = 1 To 10 Step 1
    ///         Debug.Print i
    ///     Next
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     Dim i As Long
    ///     For i = 1 To 10
    ///         Debug.Print i
    ///     Next
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    internal sealed class RemoveStepOneQuickFix : QuickFixBase
    {
        public RemoveStepOneQuickFix()
            : base(typeof(StepOneIsRedundantInspection))
        {}

        public override bool CanFixMultiple => true;
        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
        public override bool CanFixAll => true;

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.RemoveStepOneQuickFix;

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(result.QualifiedSelection.QualifiedName);
            var context = result.Context;
            rewriter.Remove(context);
        }
    }
}
