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
            IAttributeAnnotation annotation = result.Properties.Annotation;

            var attributeName = declaration.DeclarationType.HasFlag(DeclarationType.Module)
                ? annotation.Attribute
                : $"{declaration.IdentifierName}.{annotation.Attribute}";

            _attributesUpdater.AddAttribute(rewriteSession, declaration, attributeName, annotation.AttributeValues);
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.AddMissingAttributeQuickFix;

        public override CodeKind TargetCodeKind => CodeKind.AttributesCode;

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}