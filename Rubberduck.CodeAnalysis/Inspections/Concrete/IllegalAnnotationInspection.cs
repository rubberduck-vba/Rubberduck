using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.JunkDrawer.Extensions;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Extensions;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Flags invalid Rubberduck annotation comments.
    /// </summary>
    /// <why>
    /// Rubberduck is correctly parsing an annotation, but that annotation is illegal in that context.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     '@Folder("Module1.DoSomething")
    ///     Dim foo As Long
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// '@Folder("Module1.DoSomething")
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     Dim foo As Long
    /// End Sub
    /// ]]>
    /// </example>
    public sealed class IllegalAnnotationInspection : InspectionBase
    {
        public IllegalAnnotationInspection(RubberduckParserState state)
            : base(state)
        {}

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var userDeclarations = State.DeclarationFinder.AllUserDeclarations.ToList();
            var identifierReferences = State.DeclarationFinder.AllIdentifierReferences().ToList();
            var annotations = State.AllAnnotations;

            var unboundAnnotations = UnboundAnnotations(annotations, userDeclarations, identifierReferences)
                .Where(annotation => !annotation.Annotation.Target.HasFlag(AnnotationTarget.General)
                                     || annotation.AnnotatedLine == null);
            var attributeAnnotationsInDocuments = AttributeAnnotationsInDocuments(userDeclarations);

            var attributeAnnotationsOnDeclarationsNotAllowingAttributes = AttributeAnnotationsOnDeclarationsNotAllowingAttributes(userDeclarations);

            var illegalAnnotations = unboundAnnotations
                .Concat(attributeAnnotationsInDocuments)
                .Concat(attributeAnnotationsOnDeclarationsNotAllowingAttributes)
                .ToHashSet();

            return illegalAnnotations.Select(annotation => 
                new QualifiedContextInspectionResult(
                    this, 
                    string.Format(InspectionResults.IllegalAnnotationInspection, annotation.Context.annotationName().GetText()), 
                    new QualifiedContext(annotation.QualifiedSelection.QualifiedName, annotation.Context)));
        }

        private static IEnumerable<IParseTreeAnnotation> UnboundAnnotations(IEnumerable<IParseTreeAnnotation> annotations, IEnumerable<Declaration> userDeclarations, IEnumerable<IdentifierReference> identifierReferences)
        {
            var boundAnnotationsSelections = userDeclarations
                .SelectMany(declaration => declaration.Annotations)
                .Concat(identifierReferences.SelectMany(reference => reference.Annotations))
                .Select(annotation => annotation.QualifiedSelection)
                .ToHashSet();
            
            return annotations.Where(annotation => !boundAnnotationsSelections.Contains(annotation.QualifiedSelection)).ToList();
        }

        private static IEnumerable<IParseTreeAnnotation> AttributeAnnotationsInDocuments(IEnumerable<Declaration> userDeclarations)
        {
            var declarationsInDocuments = userDeclarations
                .Where(declaration => declaration.QualifiedModuleName.ComponentType == ComponentType.Document);
            return declarationsInDocuments.SelectMany(doc => doc.Annotations).Where(pta => pta.Annotation is IAttributeAnnotation);
        }

        private static IEnumerable<IParseTreeAnnotation> AttributeAnnotationsOnDeclarationsNotAllowingAttributes(IEnumerable<Declaration> userDeclarations)
        {
            return userDeclarations
                .Where(declaration => declaration.AttributesPassContext == null 
                                      && !declaration.DeclarationType.HasFlag(DeclarationType.Module))
                .SelectMany(declaration => declaration.Annotations)
                .Where(parseTreeAnnotation => parseTreeAnnotation.Annotation is IAttributeAnnotation);
        }
    }
}