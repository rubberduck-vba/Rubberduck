﻿using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Inspections.Extensions;
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
    /// <summary>
    /// Indicates that a hidden VB attribute is present for a member, but no Rubberduck annotation is documenting it.
    /// </summary>
    /// <why>
    /// Rubberduck annotations mean to document the presence of hidden VB attributes; this inspection flags members that
    /// do not have a Rubberduck annotation corresponding to the hidden VB attribute.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Public Sub DoSomething()
    /// Attribute VB_Description = "foo"
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
                // prefilter declarations to reduce searchspace
                .Where(decl => decl.QualifiedModuleName.ComponentType != ComponentType.Document
                               && !decl.IsIgnoringInspectionResultFor(AnnotationName));

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
            return Attributes.AttributeBaseName(attribute.Name, declaration.IdentifierName);
        }
    }
}