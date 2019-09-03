using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Inspections;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Indicates that the value of a hidden VB attribute is out of sync with the corresponding Rubberduck annotation comment.
    /// </summary>
    /// <why>
    /// Keeping Rubberduck annotation comments in sync with the hidden VB attribute values, surfaces these hidden attributes in the VBE code panes; 
    /// Rubberduck can rewrite the attributes to match the corresponding annotation comment.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// '@Description("foo")
    /// Public Sub DoSomething()
    /// Attribute VB_Description = "bar"
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// '@Description("foo")
    /// Public Sub DoSomething()
    /// Attribute VB_Description = "foo"
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </example>
    [CannotAnnotate]
    public sealed class AttributeValueOutOfSyncInspection : InspectionBase
    {
        public AttributeValueOutOfSyncInspection(RubberduckParserState state) 
        :base(state)
        {
        }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var declarationsWithAttributeAnnotations = State.DeclarationFinder.AllUserDeclarations
                .Where(declaration => declaration.Annotations.Any(pta => pta.Annotation is IAttributeAnnotation));
            var results = new List<DeclarationInspectionResult>();
            foreach (var declaration in declarationsWithAttributeAnnotations.Where(decl => decl.QualifiedModuleName.ComponentType != ComponentType.Document))
            {
                foreach (var annotationInstance in declaration.Annotations.Where(pta => pta.Annotation is IAttributeAnnotation))
                {
                    // cast is safe given the predicate in the foreach
                    var annotation = (IAttributeAnnotation)annotationInstance.Annotation;
                    if (HasDifferingAttributeValues(declaration, annotationInstance, out var attributeValues))
                    {
                        var attributeName = annotation.Attribute(annotationInstance);

                        var description = string.Format(InspectionResults.AttributeValueOutOfSyncInspection, 
                            attributeName, 
                            string.Join(", ", attributeValues), 
                            annotation.Name);

                        var result = new DeclarationInspectionResult(this, description, declaration,
                            new QualifiedContext(declaration.QualifiedModuleName, annotationInstance.Context));
                        result.Properties.Annotation = annotationInstance;
                        result.Properties.AttributeName = attributeName;
                        result.Properties.AttributeValues = attributeValues;

                        results.Add(result);
                    }
                }
            }

            return results;
        }

        private static bool HasDifferingAttributeValues(Declaration declaration, IParseTreeAnnotation annotationInstance, out IReadOnlyList<string> attributeValues)
        {
            if (!(annotationInstance.Annotation is IAttributeAnnotation annotation))
            {
                attributeValues = new List<string>();
                return false;
            }
            var attribute = annotation.Attribute(annotationInstance);
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
    }
}