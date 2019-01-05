using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
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