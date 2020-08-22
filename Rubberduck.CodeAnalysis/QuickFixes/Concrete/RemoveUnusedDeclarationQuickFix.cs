using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete
{
    /// <summary>
    /// Removes the declaration for a constant, variable, procedure, or line label that isn't used. This operation can break the code if the declaration is actually in use but Rubberduck couldn't find where.
    /// </summary>
    /// <inspections>
    /// <inspection name="ConstantNotUsedInspection" />
    /// <inspection name="ProcedureNotUsedInspection" />
    /// <inspection name="VariableNotUsedInspection" />
    /// <inspection name="LineLabelNotUsedInspection" />
    /// </inspections>
    /// <canfix multiple="true" procedure="false" module="true" project="true" all="true" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     Const value = 42
    ///     Debug.Print 42
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     
    ///     Debug.Print 42
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    internal sealed class RemoveUnusedDeclarationQuickFix : QuickFixBase
    {
        public RemoveUnusedDeclarationQuickFix()
            : base(typeof(ConstantNotUsedInspection), 
                  typeof(ProcedureNotUsedInspection), 
                  typeof(VariableNotUsedInspection), 
                  typeof(LineLabelNotUsedInspection))
        {}

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(result.Target.QualifiedModuleName);
            rewriter.Remove(result.Target);
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.RemoveUnusedDeclarationQuickFix;

        public override bool CanFixMultiple => true;
        public override bool CanFixInProcedure => false;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
        public override bool CanFixAll => true;
    }
}