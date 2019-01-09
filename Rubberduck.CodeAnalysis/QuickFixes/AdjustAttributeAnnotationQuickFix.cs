using System;
using System.Collections.Generic;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.Inspections.QuickFixes
{
    public class AdjustAttributeAnnotationQuickFix : QuickFixBase
    {
        private readonly IAnnotationUpdater _annotationUpdater;
        private readonly IAttributeAnnotationProvider _attributeAnnotationProvider;

        public AdjustAttributeAnnotationQuickFix(IAnnotationUpdater annotationUpdater, IAttributeAnnotationProvider attributeAnnotationProvider)
            : base(typeof(AttributeValueOutOfSyncInspection))
        {
            _annotationUpdater = annotationUpdater;
            _attributeAnnotationProvider = attributeAnnotationProvider;
        }

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            IAttributeAnnotation oldAnnotation = result.Properties.Annotation;
            string attributeName = result.Properties.AttributeName;
            IReadOnlyList<string> attributeValues = result.Properties.AttributeValues;

            var declaration = result.Target;
            if (declaration.DeclarationType.HasFlag(DeclarationType.Module))
            {
                var componentType = declaration.QualifiedModuleName.ComponentType;
                if (IsDefaultAttribute(componentType, attributeName, attributeValues))
                {
                    _annotationUpdater.RemoveAnnotation(rewriteSession, oldAnnotation);
                }
                else
                {
                    var (newAnnotationType, newAnnotationValues) = _attributeAnnotationProvider.ModuleAttributeAnnotation(attributeName, attributeValues);
                    _annotationUpdater.UpdateAnnotation(rewriteSession, oldAnnotation, newAnnotationType, newAnnotationValues);
                }
            }
            else
            {
                var attributeBaseName = AttributeBaseName(attributeName, declaration);
                var (newAnnotationType, newAnnotationValues) = _attributeAnnotationProvider.MemberAttributeAnnotation(attributeBaseName, attributeValues);
                _annotationUpdater.UpdateAnnotation(rewriteSession, oldAnnotation, newAnnotationType, newAnnotationValues);
            }
        }

        private static bool IsDefaultAttribute(ComponentType componentType, string attributeName, IReadOnlyList<string> attributeValues)
        {
            return Attributes.IsDefaultAttribute(componentType, attributeName, attributeValues);
        }

        private static string AttributeBaseName(string attributeName, Declaration declaration)
        {
            return Attributes.AttributeBaseName(attributeName, declaration.IdentifierName);
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.AdjustAttributeAnnotationQuickFix;

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}