using System.Linq;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete
{
    /// <summary>
    /// Removes a duplicated annotation comment.
    /// </summary>
    /// <inspections>
    /// <inspection name="DuplicatedAnnotationInspection" />
    /// </inspections>
    /// <canfix multiple="true" procedure="true" module="true" project="true" all="true" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// '@Obsolete
    /// '@Obsolete
    /// Public Sub DoSomething()
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// '@Obsolete
    /// Public Sub DoSomething()
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    internal sealed class RemoveDuplicatedAnnotationQuickFix : QuickFixBase
    {
        private readonly IAnnotationUpdater _annotationUpdater;

        public RemoveDuplicatedAnnotationQuickFix(IAnnotationUpdater annotationUpdater)
            : base(typeof(DuplicatedAnnotationInspection))
        {
            _annotationUpdater = annotationUpdater;
        }

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            if (!(result is IWithInspectionResultProperties<IAnnotation> resultProperties))
            {
                return;
            }

            var resultAnnotation = resultProperties.Properties;
            var duplicateAnnotations = result.Target.Annotations
                .Where(pta => pta.Annotation == resultAnnotation)
                .OrderBy(annotation => annotation.AnnotatedLine)
                .Skip(1)
                .ToList();

            _annotationUpdater.RemoveAnnotations(rewriteSession, duplicateAnnotations);
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.RemoveDuplicatedAnnotationQuickFix;

        public override bool CanFixMultiple => true;
        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
        public override bool CanFixAll => true;
    }
}
