using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Annotations.Concrete;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Refactorings.MoveToFolder
{
    public class MoveToFolderRefactoringAction : CodeOnlyRefactoringActionBase<MoveToFolderModel>
    {
        private readonly IAnnotationUpdater _annotationUpdater;

        public MoveToFolderRefactoringAction(
            IRewritingManager rewritingManager,
            IAnnotationUpdater annotationUpdater)
            : base(rewritingManager)
        {
            _annotationUpdater = annotationUpdater;
        }

        public override void Refactor(MoveToFolderModel model, IRewriteSession rewriteSession)
        {
            var oldFolderAnnotation = model.Target.Annotations.FirstOrDefault(pta => pta.Annotation is FolderAnnotation);
            if (oldFolderAnnotation == null)
            {
                AddFolderAnnotation(model, rewriteSession);
            }
            else if(!model.TargetFolder.Equals(model.Target.CustomFolder))
            {
                UpdateFolderAnnotation(model, oldFolderAnnotation, rewriteSession);
            }
        }

        private void UpdateFolderAnnotation(MoveToFolderModel model, IParseTreeAnnotation oldPta, IRewriteSession rewriteSession)
        {
            var oldFolderName = oldPta.AnnotationArguments.FirstOrDefault();
            if (oldFolderName == null || oldFolderName.Equals(model.TargetFolder))
            {
                return;
            }

            var (annotation, annotationValues) = NewAnnotation(model.TargetFolder);

            _annotationUpdater.UpdateAnnotation(rewriteSession, oldPta, annotation, annotationValues);
        }

        private static (IAnnotation annotation, IReadOnlyList<string> annotationArguments) NewAnnotation(string targetFolder)
        {
            var targetFolderLiteral = targetFolder.ToVbaStringLiteral();

            var annotation = new FolderAnnotation();
            var annotationValues = new List<string> { targetFolderLiteral };

            return (annotation, annotationValues);
        }

        private void AddFolderAnnotation(MoveToFolderModel model, IRewriteSession rewriteSession)
        {
            var (annotation, annotationValues) = NewAnnotation(model.TargetFolder);

            _annotationUpdater.AddAnnotation(rewriteSession, model.Target, annotation, annotationValues);
        }
    }
}