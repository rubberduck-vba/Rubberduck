using System.Collections.Generic;
using System.Diagnostics;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Parsing;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete
{
    /// <summary>
    /// Adjusts existing hidden attributes to match the corresponding Rubberduck annotations.
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
    /// Attribute VB_PredeclaredId = True
    /// '@PredeclaredId
    /// 
    /// Option Explicit
    /// 
    /// '@Description("Does something.")
    /// Public Sub DoSomething()
    /// Attribute VB_Description = "Does something."
    /// 
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    internal class AdjustAttributeValuesQuickFix : QuickFixBase
    {
        private readonly IAttributesUpdater _attributesUpdater;

        public AdjustAttributeValuesQuickFix(IAttributesUpdater attributesUpdater)
            : base(typeof(AttributeValueOutOfSyncInspection))
        {
            _attributesUpdater = attributesUpdater;
        }

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            if (!(result is IWithInspectionResultProperties<(IParseTreeAnnotation Annotation, string AttributeName, IReadOnlyList<string> AttributeValues)> resultProperties))
            {
                return;
            }

            var declaration = result.Target;
            var (parseTreeAnnotation, attributeBaseName, attributeValues) = resultProperties.Properties;

            var attributeName = declaration.DeclarationType.HasFlag(DeclarationType.Module)
                ? attributeBaseName
                : Attributes.MemberAttributeName(attributeBaseName, declaration.IdentifierName);

            if (!(parseTreeAnnotation.Annotation is IAttributeAnnotation attributeAnnotation))
            {
                var message = $"Tried to adjust values of attribute {attributeName} to values of non-attribute annotation {parseTreeAnnotation.Annotation.Name} in component {declaration.QualifiedModuleName}.";
                Logger.Warn(message);
                Debug.Fail(message);
                return;
            }

            var attributeValuesFromAnnotation = attributeAnnotation.AttributeValues(parseTreeAnnotation);

            _attributesUpdater.UpdateAttribute(rewriteSession, declaration, attributeName, attributeValuesFromAnnotation, oldValues: attributeValues);
        }

        public override CodeKind TargetCodeKind => CodeKind.AttributesCode;

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.AdjustAttributeValuesQuickFix;

        public override bool CanFixMultiple => true;
        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
        public override bool CanFixAll => true;
    }
}