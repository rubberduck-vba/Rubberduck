using System.Linq;
using Rubberduck.Common;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.Refactorings.Exceptions;

namespace Rubberduck.Refactorings.AnnotateDeclaration
{
    public class AnnotateDeclarationRefactoringAction : CodeOnlyRefactoringActionBase<AnnotateDeclarationModel>
    {
        private readonly IAnnotationUpdater _annotationUpdater;
        private readonly IAttributesUpdater _attributesUpdater;

        public AnnotateDeclarationRefactoringAction(
            IRewritingManager rewritingManager,
            IAnnotationUpdater annotationUpdater,
            IAttributesUpdater attributesUpdater) 
            : base(rewritingManager)
        {
            _annotationUpdater = annotationUpdater;
            _attributesUpdater = attributesUpdater;
        }

        protected override CodeKind RewriteSessionCodeKind(AnnotateDeclarationModel model)
        {
            return model.AdjustAttribute 
                   && model.Annotation is IAttributeAnnotation
                ? CodeKind.AttributesCode
                : CodeKind.CodePaneCode;
        }

        public override void Refactor(AnnotateDeclarationModel model, IRewriteSession rewriteSession)
        {
            if (model.AdjustAttribute 
                && rewriteSession.TargetCodeKind != CodeKind.AttributesCode
                && model.Annotation is IAttributeAnnotation)
            {
                throw new AttributeRewriteSessionRequiredException();
            }

            var targetDeclaration = model.Target;

            if (rewriteSession.TargetCodeKind == CodeKind.AttributesCode
                && targetDeclaration.AttributesPassContext == null
                && !targetDeclaration.DeclarationType.HasFlag(DeclarationType.Module))
            {
                throw new AttributeRewriteSessionNotSupportedException();
            }

            var arguments = model.Arguments.Select(ToCode).ToList();

            if (model.AdjustAttribute 
                && model.Annotation is IAttributeAnnotation attributeAnnotation)
            {
                var baseAttribute = attributeAnnotation.Attribute(arguments);
                var attribute = targetDeclaration.DeclarationType.HasFlag(DeclarationType.Module)
                    ? baseAttribute
                    : Attributes.MemberAttributeName(baseAttribute, targetDeclaration.IdentifierName);
                var attributeValues = attributeAnnotation.AnnotationToAttributeValues(arguments);
                _attributesUpdater.AddOrUpdateAttribute(rewriteSession, targetDeclaration, attribute, attributeValues);
            }

            _annotationUpdater.AddAnnotation(rewriteSession, targetDeclaration, model.Annotation, arguments);
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