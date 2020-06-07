using System;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Refactorings.AnnotateDeclaration
{
    public class AnnotateDeclarationRefactoringAction : CodeOnlyRefactoringActionBase<AnnotateDeclarationModel>
    {
        private readonly IAnnotationUpdater _annotationUpdater;

        public AnnotateDeclarationRefactoringAction(
            IRewritingManager rewritingManager,
            IAnnotationUpdater annotationUpdater) 
            : base(rewritingManager)
        {
            _annotationUpdater = annotationUpdater;
        }

        public override void Refactor(AnnotateDeclarationModel model, IRewriteSession rewriteSession)
        {
            var arguments = model.Arguments.Select(ToCode).ToList();
            _annotationUpdater.AddAnnotation(rewriteSession, model.Target, model.Annotation, arguments);
        }

        private string ToCode(TypedAnnotationArgument annotationArgument)
        {
            switch (annotationArgument.ArgumentType)
            {
                case AnnotationArgumentType.Text:
                    return annotationArgument.Argument.ToVbaStringLiteral();
                case AnnotationArgumentType.Boolean:
                    return ToBooleanLiteral(annotationArgument.Argument);
                default:
                    return annotationArgument.Argument;
            }
        }

        private const string NotABoolean = "NOT_A_BOOLEAN";
        private string ToBooleanLiteral(string booleanText)
        {
            if (!bool.TryParse(booleanText, out var booleanValue))
            {
                return NotABoolean;
            }

            return booleanValue
                ? Tokens.True
                : Tokens.False;
        }
    }
}