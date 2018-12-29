using System;
using System.Collections.Generic;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public class AddAttributeAnnotationQuickFix : QuickFixBase
    {
        private readonly IAnnotationUpdater _annotationUpdater;
        private readonly IAttributeAnnotationProvider _attributeAnnotationProvider;

        public AddAttributeAnnotationQuickFix(IAnnotationUpdater annotationUpdater, IAttributeAnnotationProvider attributeAnnotationProvider)
            : base(typeof(MissingModuleAnnotationInspection), typeof(MissingMemberAnnotationInspection))
        {
            _annotationUpdater = annotationUpdater;
            _attributeAnnotationProvider = attributeAnnotationProvider;
        }

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var declaration = result.Target;
            string attributeName = result.Properties.AttributeName;
            IReadOnlyList<string> attributeValues = result.Properties.AttributeValues;
            var (annotationType, annotationValues) = declaration.DeclarationType.HasFlag(DeclarationType.Module)
                ? _attributeAnnotationProvider.ModuleAttributeAnnotation(attributeName, attributeValues)
                : _attributeAnnotationProvider.MemberAttributeAnnotation(AttributeBaseName(attributeName, declaration), attributeValues);
            _annotationUpdater.AddAnnotation(rewriteSession, declaration, annotationType, annotationValues);
        }

        private static string AttributeBaseName(string attributeName, Declaration declaration)
        {
            return Attributes.AttributeBaseName(attributeName, declaration.IdentifierName);
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.AddAttributeAnnotationQuickFix;

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}