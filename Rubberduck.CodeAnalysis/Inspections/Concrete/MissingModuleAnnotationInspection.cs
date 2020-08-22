using System.Collections.Generic;
using System.Linq;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Indicates that a hidden VB attribute is present for a module, but no Rubberduck annotation is documenting it.
    /// </summary>
    /// <why>
    /// Rubberduck annotations mean to document the presence of hidden VB attributes; this inspection flags modules that
    /// do not have a Rubberduck annotation corresponding to the hidden VB attribute.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Predeclared Class">
    /// <![CDATA[
    /// Attribute VB_PredeclaredId = True
    /// Option Explicit
    /// ' ...
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Predeclared Class">
    /// <![CDATA[
    /// Attribute VB_PredeclaredId = True
    /// '@PredeclaredId
    /// Option Explicit
    /// ' ...
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class MissingModuleAnnotationInspection : DeclarationInspectionMultiResultBase<(string AttributeName, IReadOnlyList<string> AttributeValues)>
    {
        public MissingModuleAnnotationInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider, new []{DeclarationType.Module}, new []{DeclarationType.Document})
        {}

        protected override IEnumerable<(string AttributeName, IReadOnlyList<string> AttributeValues)> ResultProperties(Declaration declaration, DeclarationFinder finder)
        {
            return declaration.Attributes
                .Where(attribute => IsResultAttribute(attribute, declaration))
                .Select(PropertiesFromAttribute);
        }

        private static bool IsResultAttribute(AttributeNode attribute, Declaration declaration)
        {
            return !IsDefaultAttribute(declaration, attribute) 
                   && MissesCorrespondingModuleAnnotation(declaration, attribute);
        }

        private static (string AttributeName, IReadOnlyList<string> AttributeValues) PropertiesFromAttribute(AttributeNode attribute)
        {
            return (attribute.Name, attribute.Values);
        }

        private static bool IsDefaultAttribute(Declaration declaration, AttributeNode attribute)
        {
            return Parsing.Symbols.Attributes.IsDefaultAttribute(declaration.QualifiedModuleName.ComponentType, attribute.Name, attribute.Values);
        }

        private static bool MissesCorrespondingModuleAnnotation(Declaration declaration, AttributeNode attribute)
        {
            if (string.IsNullOrEmpty(attribute.Name) || !declaration.DeclarationType.HasFlag(DeclarationType.Module))
            {
                return false;
            }
            //VB_Ext_Key attributes are special in that identity also depends on the first value, the key.
            if (attribute.Name == "VB_Ext_Key")
            {
                return !declaration.Annotations.Where(pta => pta.Annotation is IAttributeAnnotation)
                    .Any(pta => {
                        var annotation = (IAttributeAnnotation)pta.Annotation;
                        return annotation.Attribute(pta).Equals("VB_Ext_Key") && attribute.Values[0].Equals(annotation.AttributeValues(pta)[0]);
                    });
            }

            return !declaration.Annotations.Where(pta => pta.Annotation is IAttributeAnnotation)
                .Any(pta => ((IAttributeAnnotation)pta.Annotation).Attribute(pta).Equals(attribute.Name));
        }

        protected override string ResultDescription(Declaration declaration, (string AttributeName, IReadOnlyList<string> AttributeValues) properties)
        {
            return string.Format(InspectionResults.MissingMemberAnnotationInspection,
                declaration.IdentifierName,
                properties.AttributeName,
                string.Join(", ", properties.AttributeValues));
        }
    }
}
