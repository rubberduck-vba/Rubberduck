using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class RemoveDuplicatedAnnotationQuickFix : QuickFixBase
    {
        private readonly IAnnotationUpdater _annotationUpdater;

        public RemoveDuplicatedAnnotationQuickFix(IAnnotationUpdater annotationUpdater)
            : base(typeof(DuplicatedAnnotationInspection))
        {
            _annotationUpdater = annotationUpdater;
        }

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var duplicateAnnotations = result.Target.Annotations
                .Where(annotation => annotation.AnnotationType == result.Properties.AnnotationType)
                .OrderBy(annotation => annotation.Context.Start.StartIndex)
                .Skip(1)
                .ToList();

            _annotationUpdater.RemoveAnnotations(rewriteSession, duplicateAnnotations);
        }

        public override string Description(IInspectionResult result) =>
            Resources.Inspections.QuickFixes.RemoveDuplicatedAnnotationQuickFix;

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}
