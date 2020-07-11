using System.Collections.Generic;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete
{
    /// <summary>
    /// Adjusts existing Rubberduck annotations to match the corresponding hidden attributes.
    /// </summary>
    /// <inspections>
    /// <inspection name="AttributeValueOutOfSyncInspection" />
    /// </inspections>
    /// <canfix multiple="true" procedure="true" module="true" project="true" all="true" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Attribute VB_PredeclaredId = False
    /// '@PredeclaredId
    /// 
    /// Option Explicit
    /// 
    /// '@Description("Does something.")
    /// Public Sub DoSomething()
    /// Attribute VB_Description = "Does something else."
    /// 
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// 
    /// Option Explicit
    /// 
    /// '@Description("Does something else.")
    /// Public Sub DoSomething()
    /// Attribute VB_Description = "Does something else."
    /// 
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    internal class AdjustAttributeAnnotationQuickFix : QuickFixBase
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
            if (!(result is IWithInspectionResultProperties<(IParseTreeAnnotation Annotation, string AttributeName, IReadOnlyList<string> AttributeValues)> resultProperties))
            {
                return;
            }

            var declaration = result.Target;
            var (oldParseTreeAnnotation, attributeBaseName, attributeValues) = resultProperties.Properties;

            if (declaration.DeclarationType.HasFlag(DeclarationType.Module))
            {
                var componentType = declaration.QualifiedModuleName.ComponentType;
                if (IsDefaultAttribute(componentType, attributeBaseName, attributeValues))
                {
                    _annotationUpdater.RemoveAnnotation(rewriteSession, oldParseTreeAnnotation);
                }
                else
                {
                    var (newAnnotation, newAnnotationValues) = _attributeAnnotationProvider.ModuleAttributeAnnotation(attributeBaseName, attributeValues);
                    _annotationUpdater.UpdateAnnotation(rewriteSession, oldParseTreeAnnotation, newAnnotation, newAnnotationValues);
                }
            }
            else
            {
                var (newAnnotation, newAnnotationValues) = _attributeAnnotationProvider.MemberAttributeAnnotation(attributeBaseName, attributeValues);
                _annotationUpdater.UpdateAnnotation(rewriteSession, oldParseTreeAnnotation, newAnnotation, newAnnotationValues);
            }
        }

        private static bool IsDefaultAttribute(ComponentType componentType, string attributeName, IReadOnlyList<string> attributeValues)
        {
            return Attributes.IsDefaultAttribute(componentType, attributeName, attributeValues);
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.AdjustAttributeAnnotationQuickFix;

        public override bool CanFixMultiple => true;
        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
        public override bool CanFixAll => true;
    }
}