using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete
{
    /// <summary>
    /// Removes 'Option Base 0' statement from a module, making it implicit (0 being the default implicit lower bound for implicitly-sized arrays).
    /// </summary>
    /// <inspections>
    /// <inspection name="RedundantOptionInspection" />
    /// </inspections>
    /// <canfix multiple="true" procedure="false" module="false" project="false" all="false" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Option Explicit
    /// Option Base 0
    /// 
    /// Public Sub DoSomething()
    ///     Dim values(10) ' implicit lower bound is 0
    ///     Debug.Print LBound(values), UBound(values)
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     Dim values(10) ' implicit lower bound is 0
    ///     Debug.Print LBound(values), UBound(values)
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    internal sealed class RemoveRedundantOptionStatementQuickFix : QuickFixBase
    {
        public RemoveRedundantOptionStatementQuickFix()
            : base(typeof(RedundantOptionInspection))
        {}

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(result.QualifiedSelection.QualifiedName);
            rewriter.Remove(result.Context);
        }

        public override string Description(IInspectionResult result)
        {
            return string.Format(Resources.Inspections.QuickFixes.RemoveRedundantOptionStatementQuickFix, result.Context.GetText());
        }

        public override bool CanFixMultiple => true;
        public override bool CanFixInProcedure => false;
        public override bool CanFixInModule => false;
        public override bool CanFixInProject => false;
        public override bool CanFixAll => false;
    }
}