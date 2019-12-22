using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Inspections.Extensions;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Indicates that a hidden VB attribute is present for a module, but no Rubberduck annotation is documenting it.
    /// </summary>
    /// <why>
    /// Rubberduck annotations mean to document the presence of hidden VB attributes; this inspection flags modules that
    /// do not have a Rubberduck annotation corresponding to the hidden VB attribute.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Attribute VB_PredeclaredId = True
    /// Option Explicit
    /// ' ...
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Attribute VB_PredeclaredId = True
    /// '@PredeclaredId
    /// Option Explicit
    /// ' ...
    /// ]]>
    /// </example>
    public sealed class MissingModuleAnnotationInspection : InspectionBase
    {
        public MissingModuleAnnotationInspection(RubberduckParserState state) 
        :base(state)
        {}

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var moduleDeclarationsWithAttributes = State.DeclarationFinder
                .UserDeclarations(DeclarationType.Module)
                .Where(decl => decl.Attributes.Any());

            var declarationsToInspect = moduleDeclarationsWithAttributes
                // prefilter declarations to reduce searchspace
                .Where(decl => decl.QualifiedModuleName.ComponentType != ComponentType.Document
                               && !decl.IsIgnoringInspectionResultFor(AnnotationName));

            var results = new List<DeclarationInspectionResult>();
            foreach (var declaration in declarationsToInspect)
            {
                foreach (var attribute in declaration.Attributes)
                {
                    if (IsDefaultAttribute(declaration, attribute))
                    {
                        continue;
                    }

                    if (MissesCorrespondingModuleAnnotation(declaration, attribute))
                    {
                        var description = string.Format(InspectionResults.MissingMemberAnnotationInspection,
                            declaration.IdentifierName,
                            attribute.Name,
                            string.Join(", ", attribute.Values));

                        var result = new DeclarationInspectionResult(this, description, declaration,
                            new QualifiedContext(declaration.QualifiedModuleName, declaration.Context));
                        result.Properties.AttributeName = attribute.Name;
                        result.Properties.AttributeValues = attribute.Values;

                        results.Add(result);
                    }
                }
            }

            return results;
        }

        private static bool IsDefaultAttribute(Declaration declaration, AttributeNode attribute)
        {
            return Attributes.IsDefaultAttribute(declaration.QualifiedModuleName.ComponentType, attribute.Name, attribute.Values);
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
    }
}