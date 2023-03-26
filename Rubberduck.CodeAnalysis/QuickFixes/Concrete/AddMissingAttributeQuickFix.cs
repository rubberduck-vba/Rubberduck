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
    /// Exports the module, adds the hidden attributes as needed, re-imports the temporary file back into the project.
    /// </summary>
    /// <inspections>
    /// <inspection name="MissingAttributeInspection" />
    /// </inspections>
    /// <canfix multiple="true" procedure="true" module="true" project="true" all="true" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// '@ModuleDescription("Just a module.")
    /// Option Explicit
    /// 
    /// '@Description("Does something.")
    /// Public Sub DoSomething()
    /// 
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Attribute VB_Description = "Just a module."
    /// '@ModuleDescription("Just a module.")
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
    internal sealed class AddMissingAttributeQuickFix : QuickFixBase
    {
        private readonly IAttributesUpdater _attributesUpdater; 

        public AddMissingAttributeQuickFix(IAttributesUpdater attributesUpdater)
            : base(typeof(MissingAttributeInspection))
        {
            _attributesUpdater = attributesUpdater;
        }

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            if (!(result is IWithInspectionResultProperties<IParseTreeAnnotation> resultProperties))
            {
                return;
            }

            var declaration = result.Target;
            var annotationInstance = resultProperties.Properties;
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

        public override CodeKind TargetCodeKind => CodeKind.AttributesCode;

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.AddMissingAttributeQuickFix;
        
        public override bool CanFixMultiple => true;
        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
        public override bool CanFixAll => true;
    }
}