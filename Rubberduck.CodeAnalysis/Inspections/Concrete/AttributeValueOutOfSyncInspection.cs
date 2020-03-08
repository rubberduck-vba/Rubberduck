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
    /// Indicates that the value of a hidden VB attribute is out of sync with the corresponding Rubberduck annotation comment.
    /// </summary>
    /// <why>
    /// Keeping Rubberduck annotation comments in sync with the hidden VB attribute values, surfaces these hidden attributes in the VBE code panes; 
    /// Rubberduck can rewrite the attributes to match the corresponding annotation comment.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// '@Description("foo")
    /// Public Sub DoSomething()
    /// Attribute VB_Description = "bar"
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
    internal sealed class AttributeValueOutOfSyncInspection : DeclarationInspectionMultiResultBase<(IParseTreeAnnotation Annotation, string AttributeName, IReadOnlyList<string> AttributeValues)>
    {
        public AttributeValueOutOfSyncInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {}

        protected override IEnumerable<(IParseTreeAnnotation Annotation, string AttributeName, IReadOnlyList<string> AttributeValues)> ResultProperties(Declaration declaration, DeclarationFinder finder)
        {
            if (declaration.QualifiedModuleName.ComponentType == ComponentType.Document)
            {
                return Enumerable.Empty<(IParseTreeAnnotation Annotation, string AttributeName, IReadOnlyList<string> AttributeValues)>();
            }

            return OutOfSyncAttributeAnnotations(declaration);
        }

        private static IEnumerable<(IParseTreeAnnotation Annotation, string AttributeName, IReadOnlyList<string> AttributeValues)> OutOfSyncAttributeAnnotations(Declaration declaration)
        {
            foreach (var pta in declaration.Annotations)
            {
                if (!(pta.Annotation is IAttributeAnnotation annotation) 
                    || !HasDifferingAttributeValues(declaration, pta, out var attributeValues))
                {
                    continue;
                }

                var attributeName = annotation.Attribute(pta);
                yield return (pta, attributeName, attributeValues);
            }
        }

        private static bool HasDifferingAttributeValues(Declaration declaration, IParseTreeAnnotation annotationInstance, out IReadOnlyList<string> attributeValues)
        {
            if (!(annotationInstance.Annotation is IAttributeAnnotation annotation))
            {
                attributeValues = new List<string>();
                return false;
            }

            var attributeNodes = declaration.DeclarationType.HasFlag(DeclarationType.Module)
                ? declaration.Attributes.AttributeNodesFor(annotationInstance)
                : declaration.Attributes.AttributeNodesFor(annotationInstance, declaration.IdentifierName);

            foreach (var attributeNode in attributeNodes)
            {
                var values = attributeNode.Values;
                if (!annotation.AttributeValues(annotationInstance).SequenceEqual(values))
                {
                    attributeValues = values;
                    return true;
                }
            }
            attributeValues = new List<string>();
            return false;
        }

        protected override string ResultDescription(Declaration declaration, (IParseTreeAnnotation Annotation, string AttributeName, IReadOnlyList<string> AttributeValues) properties)
        {
            var (pta, attributeName, attributeValues) = properties;
            var annotationName = pta.Annotation.Name;
            return string.Format(InspectionResults.AttributeValueOutOfSyncInspection,
                attributeName,
                string.Join(", ", attributeValues),
                annotationName);
        }
    }
}