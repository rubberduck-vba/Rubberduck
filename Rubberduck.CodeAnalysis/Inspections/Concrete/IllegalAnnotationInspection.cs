using System.Collections.Generic;
using System.Linq;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.CodeAnalysis.Inspections.Results;
using Rubberduck.JunkDrawer.Extensions;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Flags invalid Rubberduck annotation comments.
    /// </summary>
    /// <why>
    /// Rubberduck is correctly parsing an annotation, but that annotation is illegal in that context.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     '@Folder("Module1.DoSomething")
    ///     Dim foo As Long
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// '@Folder("Module1.DoSomething")
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     Dim foo As Long
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class IllegalAnnotationInspection : InspectionBase
    {
        public IllegalAnnotationInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {}

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults(DeclarationFinder finder)
        {
            return finder.UserDeclarations(DeclarationType.Module)
                .Where(module => module != null)
                .SelectMany(module => DoGetInspectionResults(module.QualifiedModuleName, finder));
        }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module, DeclarationFinder finder)
        {
            var userDeclarations = finder.Members(module).ToList();
            var identifierReferences = finder.IdentifierReferences(module).ToList();
            var annotations = finder.FindAnnotations(module);

            var unboundAnnotations = UnboundAnnotations(annotations, userDeclarations, identifierReferences)
                .Where(annotation => !annotation.Annotation.Target.HasFlag(AnnotationTarget.General)
                                     || annotation.AnnotatedLine == null);

            var attributeAnnotationsOnDeclarationsNotAllowingAttributes = AttributeAnnotationsOnDeclarationsNotAllowingAttributes(userDeclarations);

            var illegalAnnotations = unboundAnnotations
                .Concat(attributeAnnotationsOnDeclarationsNotAllowingAttributes)
                .Distinct();

            if (module.ComponentType == ComponentType.Document)
            {
                var attributeAnnotationsInDocuments = AttributeAnnotationsInDocuments(userDeclarations);
                illegalAnnotations = illegalAnnotations
                    .Concat(attributeAnnotationsInDocuments)
                    .Distinct();
            }

            return illegalAnnotations
                .Select(InspectionResult)
                .ToList();
        }

        private IInspectionResult InspectionResult(IParseTreeAnnotation pta)
        {
            return new QualifiedContextInspectionResult(
                this,
                ResultDescription(pta),
                Context(pta));
        }

        private static string ResultDescription(IParseTreeAnnotation pta)
        {
            var annotationText = pta.Context.annotationName().GetText();
            return string.Format(
                InspectionResults.IllegalAnnotationInspection,
                annotationText);
        }

        private static QualifiedContext Context(IParseTreeAnnotation pta)
        {
            return new QualifiedContext(pta.QualifiedSelection.QualifiedName, pta.Context);
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
            return declarationsInDocuments.SelectMany(doc => doc.Annotations)
                .Where(pta => pta.Annotation is IAttributeAnnotation);
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