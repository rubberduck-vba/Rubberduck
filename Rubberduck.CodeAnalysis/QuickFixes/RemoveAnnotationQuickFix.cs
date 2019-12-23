using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    /// <summary>
    /// Removes an annotation comment representing a hidden module or member attribute, in order to maintain consistency between hidden attributes and annotation comments.
    /// </summary>
    /// <inspections>
    /// <inspection name="MissingAttributeInspection" />
    /// </inspections>
    /// <canfix procedure="false" module="false" project="false" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Attribute VB_PredeclaredId = False
    /// '@PredeclaredId
    /// 
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Attribute VB_PredeclaredId = False
    /// 
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    public sealed class RemoveAnnotationQuickFix : QuickFixBase
    {
        private readonly IAnnotationUpdater _annotationUpdater; 

        public RemoveAnnotationQuickFix(IAnnotationUpdater annotationUpdater)
        :base(typeof(MissingAttributeInspection))
        {
            _annotationUpdater = annotationUpdater;
        }

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            _annotationUpdater.RemoveAnnotation(rewriteSession, result.Properties.Annotation);
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.RemoveAnnotationQuickFix;

        public override bool CanFixInProcedure => false;
        public override bool CanFixInModule => false;
        public override bool CanFixInProject => false;
    }
}