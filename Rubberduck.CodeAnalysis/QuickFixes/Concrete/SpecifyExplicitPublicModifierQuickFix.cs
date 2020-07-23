using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete
{
    /// <summary>
    /// Introduces an explicit 'Public' access modifier for a procedure that is implicitly public.
    /// </summary>
    /// <inspections>
    /// <inspection name="ImplicitPublicMemberInspection" />
    /// </inspections>
    /// <canfix multiple="true" procedure="false" module="true" project="true" all="true" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Sub DoSomething()
    ///     Debug.Print 42
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     Debug.Print 42
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    internal sealed class SpecifyExplicitPublicModifierQuickFix : QuickFixBase
    {
        public SpecifyExplicitPublicModifierQuickFix()
            : base(typeof(ImplicitPublicMemberInspection))
        {}

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(result.Target.QualifiedModuleName);
            rewriter.InsertBefore(result.Context.Start.TokenIndex, "Public ");
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.SpecifyExplicitPublicModifierQuickFix;

        public override bool CanFixMultiple => true;
        public override bool CanFixInProcedure => false;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
        public override bool CanFixAll => true;
    }
}