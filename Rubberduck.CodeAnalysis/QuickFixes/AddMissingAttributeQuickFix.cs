using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Parsing;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class AddMissingAttributeQuickFix : QuickFixBase
    {
        private readonly IAttributesUpdater _attributesUpdater; 

        public AddMissingAttributeQuickFix(IAttributesUpdater attributesUpdater)
            : base(typeof(MissingAttributeInspection))
        {
            _attributesUpdater = attributesUpdater;
        }

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var declaration = result.Target;
            IParseTreeAnnotation annotationInstance = result.Properties.Annotation;
            if (!(annotationInstance.Annotation is IAttributeAnnotation annotation))
            {
                return;
            }
            var attribute = annotation.Attribute(annotationInstance);
            var attributeName = declaration.DeclarationType.HasFlag(DeclarationType.Module)
                ? attribute
                : $"{declaration.IdentifierName}.{attribute}";

            _attributesUpdater.AddAttribute(rewriteSession, declaration, attributeName, annotation.AttributeValues(annotationInstance));
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.AddMissingAttributeQuickFix;

        public override CodeKind TargetCodeKind => CodeKind.AttributesCode;

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}