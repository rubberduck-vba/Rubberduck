using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Inspections.Extensions;
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
    /// Indicates that a Rubberduck annotation is documenting the presence of a VB attribute that is actually missing.
    /// </summary>
    /// <why>
    /// Rubberduck annotations mean to document the presence of hidden VB attributes; this inspection flags annotations that
    /// do not have a corresponding VB attribute.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// '@Description("foo")
    /// Public Sub DoSomething()
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
    public sealed class MissingAttributeInspection : InspectionBase
    {
        public MissingAttributeInspection(RubberduckParserState state)
            : base(state)
        {}

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var declarationsWithAttributeAnnotations = State.DeclarationFinder.AllUserDeclarations
                .Where(declaration => declaration.Annotations.Any(pta => pta.Annotation is IAttributeAnnotation)
                    && (declaration.DeclarationType.HasFlag(DeclarationType.Module) 
                        || declaration.AttributesPassContext != null));
            var results = new List<DeclarationInspectionResult>();

            // prefilter declarations to reduce searchspace
            var interestingDeclarations = declarationsWithAttributeAnnotations.Where(decl => decl.QualifiedModuleName.ComponentType != ComponentType.Document
                                                                                                   && !decl.IsIgnoringInspectionResultFor(AnnotationName));
            foreach (var declaration in interestingDeclarations)
            {
                foreach (var annotationInstance in declaration.Annotations.Where(pta => pta.Annotation is IAttributeAnnotation))
                {
                    var annotation = (IAttributeAnnotation)annotationInstance.Annotation;
                    if (MissesCorrespondingAttribute(declaration, annotationInstance))
                    {
                        var description = string.Format(InspectionResults.MissingAttributeInspection, declaration.IdentifierName, annotation.Name);

                        var result = new DeclarationInspectionResult(this, description, declaration,
                            new QualifiedContext(declaration.QualifiedModuleName, annotationInstance.Context));
                        result.Properties.Annotation = annotationInstance;

                        results.Add(result);
                    }
                }
            }

            return results;
        }

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