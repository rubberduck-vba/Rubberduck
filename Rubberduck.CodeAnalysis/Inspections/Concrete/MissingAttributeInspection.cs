using System.Collections.Generic;
using System.Linq;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.CodeAnalysis.Inspections.Attributes;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Indicates that a Rubberduck annotation is documenting the presence of a VB attribute that is actually missing.
    /// </summary>
    /// <why>
    /// Rubberduck annotations mean to document the presence of hidden VB attributes; this inspection flags annotations that
    /// do not have a corresponding VB attribute.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// '@Description("foo")
    /// Public Sub DoSomething()
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// '@Description("foo")
    /// Public Sub DoSomething()
    /// Attribute VB_Description = "foo"
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    [CannotAnnotate]
    internal sealed class MissingAttributeInspection : DeclarationInspectionMultiResultBase<IParseTreeAnnotation>
    {
        public MissingAttributeInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {}

        protected override IEnumerable<IParseTreeAnnotation> ResultProperties(Declaration declaration, DeclarationFinder finder)
        {
            if (!declaration.DeclarationType.HasFlag(DeclarationType.Module)
                && declaration.AttributesPassContext == null
                || declaration.QualifiedModuleName.ComponentType == ComponentType.Document)
            {
                return Enumerable.Empty<IParseTreeAnnotation>();
            }

            return declaration.Annotations
                .Where(pta => pta.Annotation is IAttributeAnnotation && MissesCorrespondingAttribute(declaration, pta));
        }

        protected override string ResultDescription(Declaration declaration, IParseTreeAnnotation pta) =>
            string.Format(InspectionResults.MissingAttributeInspection, declaration.IdentifierName, pta.Annotation.Name);
        

        private static bool MissesCorrespondingAttribute(Declaration declaration, IParseTreeAnnotation annotationInstance)
        {
            if (!(annotationInstance.Annotation is IAttributeAnnotation annotation))
            {
                return false;
            }
            var attribute = annotation.Attribute(annotationInstance);
            if (string.IsNullOrEmpty(attribute))
            {
                return false;
            }
            return declaration.DeclarationType.HasFlag(DeclarationType.Module)
                ? !declaration.Attributes.HasAttributeFor(annotationInstance)
                : !declaration.Attributes.HasAttributeFor(annotationInstance, declaration.IdentifierName);
        }
    }
}