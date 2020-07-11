using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete
{
    /// <summary>
    /// Removes an annotation comment representing a hidden module or member attribute, in order to maintain consistency between hidden attributes and annotation comments.
    /// </summary>
    /// <inspections>
    /// <inspection name="MissingAttributeInspection" />
    /// </inspections>
    /// <canfix multiple="true" procedure="false" module="false" project="false" all="false" />
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
    internal sealed class RemoveAnnotationQuickFix : QuickFixBase
    {
        private readonly IAnnotationUpdater _annotationUpdater; 

        public RemoveAnnotationQuickFix(IAnnotationUpdater annotationUpdater)
        :base(typeof(MissingAttributeInspection))
        {
            _annotationUpdater = annotationUpdater;
        }

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            if (!(result is IWithInspectionResultProperties<IParseTreeAnnotation> resultProperties))
            {
                return;
            }

            _annotationUpdater.RemoveAnnotation(rewriteSession, resultProperties.Properties);
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.RemoveAnnotationQuickFix;

        public override bool CanFixMultiple => true;
        public override bool CanFixInProcedure => false;
        public override bool CanFixInModule => false;
        public override bool CanFixInProject => false;
        public override bool CanFixAll => false;
    }
}