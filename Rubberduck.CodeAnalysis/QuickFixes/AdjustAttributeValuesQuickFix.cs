using System;
using System.Collections.Generic;
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
    public class AdjustAttributeValuesQuickFix : QuickFixBase
    {
        private readonly IAttributesUpdater _attributesUpdater;

        public AdjustAttributeValuesQuickFix(IAttributesUpdater attributesUpdater)
            : base(typeof(AttributeValueOutOfSyncInspection))
        {
            _attributesUpdater = attributesUpdater;
        }

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var declaration = result.Target;
            IAttributeAnnotation annotation = result.Properties.Annotation;
            IReadOnlyList<string> attributeValues = result.Properties.AttributeValues;

            var attributeName = declaration.DeclarationType.HasFlag(DeclarationType.Module)
                ? annotation.Attribute
                : $"{declaration.IdentifierName}.{annotation.Attribute}";

            _attributesUpdater.UpdateAttribute(rewriteSession, declaration, attributeName, annotation.AttributeValues, attributeValues);
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.AdjustAttributeValuesQuickFix;

        public override CodeKind TargetCodeKind => CodeKind.AttributesCode;

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}