using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class MissingMemberAnnotationInspection : InspectionBase
    {
        public MissingMemberAnnotationInspection(RubberduckParserState state) 
        :base(state)
        {}

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var memberDeclarationsWithAttributes = State.DeclarationFinder.AllUserDeclarations
                .Where(decl => !decl.DeclarationType.HasFlag(DeclarationType.Module)
                                && decl.Attributes.Any());

            var declarationsToInspect = memberDeclarationsWithAttributes
                .Where(decl => decl.QualifiedModuleName.ComponentType != ComponentType.Document
                               && !IsIgnoringInspectionResultFor(decl, AnnotationName));

            var results = new List<DeclarationInspectionResult>();
            foreach (var declaration in declarationsToInspect)
            {
                foreach (var attribute in declaration.Attributes)
                {
                    if (MissesCorrespondingMemberAnnotation(declaration, attribute))
                    {
                        var attributeBaseName = AttributeBaseName(declaration, attribute);

                        var description = string.Format(InspectionResults.MissingMemberAnnotationInspection, 
                            declaration.IdentifierName,
                            attributeBaseName,
                            string.Join(", ", attribute.Values));

                        var result = new DeclarationInspectionResult(this, description, declaration,
                            new QualifiedContext(declaration.QualifiedModuleName, declaration.Context));
                        result.Properties.AttributeName = attributeBaseName;
                        result.Properties.AttributeValues = attribute.Values;

                        results.Add(result);
                    }
                }
            }

            return results;
        }

        private static bool MissesCorrespondingMemberAnnotation(Declaration declaration, AttributeNode attribute)
        {
            if (string.IsNullOrEmpty(attribute.Name) || declaration.DeclarationType.HasFlag(DeclarationType.Module))
            {
                return false;
            }

            var attributeBaseName = AttributeBaseName(declaration, attribute);

            //VB_Ext_Key attributes are special in that identity also depends on the first value, the key.
            if (attributeBaseName == "VB_Ext_Key")
            {
                return !declaration.Annotations.OfType<IAttributeAnnotation>()
                    .Any(annotation => annotation.Attribute.Equals("VB_Ext_Key") && attribute.Values[0].Equals(annotation.AttributeValues[0]));
            }

            return !declaration.Annotations.OfType<IAttributeAnnotation>()
                .Any(annotation => annotation.Attribute.Equals(attributeBaseName));
        }

        private static string AttributeBaseName(Declaration declaration, AttributeNode attribute)
        {
            return Attributes.AttributeBaseName(attribute.Name, declaration.IdentifierName);
        }
    }
}