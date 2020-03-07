using System.Linq;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    /// <summary>
    /// Removes a duplicated annotation comment.
    /// </summary>
    /// <inspections>
    /// <inspection name="DuplicatedAnnotationInspection" />
    /// </inspections>
    /// <canfix procedure="true" module="true" project="true" />
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

        public override string Description(IInspectionResult result) =>
            Resources.Inspections.QuickFixes.RemoveDuplicatedAnnotationQuickFix;

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}
