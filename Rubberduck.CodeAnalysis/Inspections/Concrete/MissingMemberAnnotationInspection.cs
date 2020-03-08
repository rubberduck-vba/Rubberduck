using System.Collections.Generic;
using System.Linq;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Indicates that a hidden VB attribute is present for a member, but no Rubberduck annotation is documenting it.
    /// </summary>
    /// <why>
    /// Rubberduck annotations mean to document the presence of hidden VB attributes; this inspection flags members that
    /// do not have a Rubberduck annotation corresponding to the hidden VB attribute.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    /// Attribute VB_Description = "foo"
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
    internal sealed class MissingMemberAnnotationInspection : DeclarationInspectionMultiResultBase<(string AttributeName, IReadOnlyList<string> AttriguteValues)>
    {
        public MissingMemberAnnotationInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider, new DeclarationType[0], new []{DeclarationType.Module })
        {}

        protected override IEnumerable<(string AttributeName, IReadOnlyList<string> AttriguteValues)> ResultProperties(Declaration declaration, DeclarationFinder finder)
        {
            if (declaration.QualifiedModuleName.ComponentType == ComponentType.Document)
            {
                return Enumerable.Empty<(string AttributeName, IReadOnlyList<string> AttriguteValues)>();
            }

            return declaration.Attributes
                .Where(attribute => MissesCorrespondingMemberAnnotation(declaration, attribute))
                .Select(attribute => (AttributeBaseName(declaration, attribute), attribute.Values));
        }

        private static bool MissesCorrespondingMemberAnnotation(Declaration declaration, AttributeNode attribute)
        {
            if (string.IsNullOrEmpty(attribute.Name) || declaration.DeclarationType.HasFlag(DeclarationType.Module))
            {
                return false;
            }

            var attributeBaseName = AttributeBaseName(declaration, attribute);
            // VB_Ext_Key attributes are special in that identity also depends on the first value, the key.
            if (attributeBaseName == "VB_Ext_Key")
            {
                return !declaration.Annotations.Where(pta => pta.Annotation is IAttributeAnnotation)
                    .Any(pta => {
                            var annotation = (IAttributeAnnotation)pta.Annotation;
                            return annotation.Attribute(pta).Equals("VB_Ext_Key") && attribute.Values[0].Equals(annotation.AttributeValues(pta)[0]);
                        });
            }

            return !declaration.Annotations.Where(pta => pta.Annotation is IAttributeAnnotation)
                .Any(pta => ((IAttributeAnnotation)pta.Annotation).Attribute(pta).Equals(attributeBaseName));
        }

        private static string AttributeBaseName(Declaration declaration, AttributeNode attribute)
        {
            return Parsing.Symbols.Attributes.AttributeBaseName(attribute.Name, declaration.IdentifierName);
        }

        protected override string ResultDescription(Declaration declaration, (string AttributeName, IReadOnlyList<string> AttriguteValues) properties)
        {
            var (attributeBaseName, attributeValues) = properties;
            return string.Format(InspectionResults.MissingMemberAnnotationInspection,
                declaration.IdentifierName,
                attributeBaseName,
                string.Join(", ", attributeValues));
        }
    }
}